"""Analysis File Builder - upload any CSV / Excel files, get one verified analysis-ready table.
Run locally:  streamlit run app.py
"""
import base64, datetime as dt, hmac, io, os, shutil, tempfile, zipfile
import pandas as pd
import streamlit as st

from branding import BRAND
from pipeline.generic import PlanError, base_options, execute_plan, load_tables, make_plan, table_label, update_item
from pipeline.generic_exports import (DMETA, build_output, helti_applicable, helti_derived, pii_columns,
                                      to_csv, to_excel, to_package, to_sav)

HERE = os.path.dirname(os.path.abspath(__file__))
B = BRAND
st.set_page_config(page_title=f"{B['app_title']} | {B['name']}", page_icon="📊", layout="wide")
SS = st.session_state


# ---------------------------------------------------------------- styling
def logo_html(height=44):
    for name, mime in (("logo.png", "image/png"), ("logo.svg", "image/svg+xml"), ("logo_placeholder.svg", "image/svg+xml")):
        p = os.path.join(HERE, "assets", name)
        if os.path.exists(p):
            b64 = base64.b64encode(open(p, "rb").read()).decode()
            return f'<img src="data:{mime};base64,{b64}" alt="{B["name"]} logo" style="height:{height}px">'
    return f'<span style="font-weight:700;font-size:1.3rem;color:{B["PRIMARY"]}">{B["name"]}</span>'


st.markdown(f"""
<style>
  :root {{ --brand:{B['PRIMARY']}; --accent:{B['ACCENT']}; --surface:{B['SURFACE']}; --muted:{B['MUTED']}; }}
  .block-container {{ padding-top: 1.2rem; max-width: 1200px; }}
  header[data-testid="stHeader"] {{ background: transparent; }}
  .ai-topbar {{ display:flex; align-items:center; justify-content:space-between; gap:1rem; flex-wrap:wrap;
               padding:.6rem 0 1rem; border-bottom:1px solid #E3E9F0; margin-bottom:1.2rem; }}
  .ai-topbar a {{ color:var(--brand); text-decoration:none; font-weight:600; font-size:.9rem;
                 border:1px solid var(--brand); border-radius:999px; padding:.35rem .9rem; }}
  .ai-topbar a:hover {{ background:var(--brand); color:#fff; }}
  .ai-hero h1 {{ font-size:1.9rem; line-height:1.2; margin:0 0 .3rem; color:var(--brand); }}
  .ai-hero p {{ color:var(--muted); margin:0 0 1.2rem; font-size:1.02rem; max-width:52rem; }}
  .ai-step {{ display:inline-flex; align-items:center; gap:.5rem; font-weight:700; color:var(--brand);
             font-size:1.05rem; margin:1.2rem 0 .4rem; }}
  .ai-step span {{ background:var(--accent); color:#fff; border-radius:999px; width:1.6rem; height:1.6rem;
                  display:inline-flex; align-items:center; justify-content:center; font-size:.85rem; }}
  .ai-row {{ font-size:.92rem; line-height:1.35; }}
  .ai-row small {{ color:var(--muted); }}
  div[data-testid="stMetric"] {{ background:var(--surface); border:1px solid #E3E9F0; border-radius:12px; padding:.7rem .9rem; }}
  div[data-testid="stMetricValue"] {{ color:var(--brand); }}
  .stButton>button[kind="primary"], .stDownloadButton>button {{ border-radius:10px; font-weight:600; }}
  .stDownloadButton>button {{ border:1px solid var(--brand); color:var(--brand); width:100%; }}
  .stProgress > div > div > div > div {{ background-color: var(--accent); }}
  .ai-footer {{ margin-top:2.5rem; padding:1rem 0 .5rem; border-top:1px solid #E3E9F0; color:var(--muted);
               font-size:.85rem; display:flex; justify-content:space-between; flex-wrap:wrap; gap:.5rem; }}
  .ai-footer a {{ color:var(--brand); font-weight:600; text-decoration:none; }}
</style>
<div class="ai-topbar"><div>{logo_html()}</div>
  <a href="{B['url']}" target="_blank" rel="noopener">{B['domain']}</a></div>
<div class="ai-hero"><h1>{B['app_title']}</h1><p>{B['app_subtitle']}</p></div>
""", unsafe_allow_html=True)


def footer():
    st.markdown(f"""<div class="ai-footer"><div>{B['powered_by']}</div>
      <div><a href="{B['url']}" target="_blank" rel="noopener">{B['domain']}</a> &middot; {B['tagline']}</div></div>""",
                unsafe_allow_html=True)


def step(n, text):
    st.markdown(f'<div class="ai-step"><span>{n}</span> {text}</div>', unsafe_allow_html=True)


# ---------------------------------------------------------------- optional password gate
def password_ok():
    try:
        secret = st.secrets.get("APP_PASSWORD")
    except Exception:
        secret = None
    if not secret or SS.get("auth_ok"):
        return True
    with st.form("login"):
        st.markdown("**This workspace is private.** Enter the access password.")
        pw = st.text_input("Password", type="password")
        if st.form_submit_button("Enter", type="primary"):
            if hmac.compare_digest(pw, str(secret)):
                SS.auth_ok = True; st.rerun()
            st.error("Incorrect password.")
    return False


if not password_ok():
    footer(); st.stop()

# ---------------------------------------------------------------- sidebar
with st.sidebar:
    st.markdown(logo_html(38), unsafe_allow_html=True)
    st.markdown("### How it works")
    st.markdown(
        "1. **Upload** CSV / Excel files (or one .zip).\n"
        "2. **Study**: every sheet is profiled - header row, column types, empty rows, likely ID columns.\n"
        "3. **Plan**: the builder picks the base table and the ID linking each file, and how to handle "
        "files with several rows per ID. You can change any choice.\n"
        "4. **Build** one row per ID, with a full log of every decision.\n"
        "5. **Download** CSV, Excel, SPSS .sav or an SPSS package.")
    st.markdown("### Rules it follows")
    st.caption("IDs written differently (001001 / 1001, ' p-01 ' / P-01) are matched. Byte-identical or same-content "
               "files are skipped. A column repeated in several files with the same values is kept once; different "
               "values are kept with the file name added. Copies of base-table text columns (e.g. names) are dropped "
               "and every mismatch is listed.")
    st.markdown("### Data privacy")
    st.caption("Files are processed in a temporary folder for this session only and deleted with **Clear session** or "
               "when the session ends. Nothing is stored or shared.")
    if st.button("Clear session", use_container_width=True):
        if SS.get("workdir"):
            shutil.rmtree(SS.workdir, ignore_errors=True)
        SS.clear(); st.rerun()

if "workdir" not in SS:
    SS.workdir = tempfile.mkdtemp(prefix="afb_")
WORK = SS.workdir
IN_DIR = os.path.join(WORK, "input")
EXTS = (".csv", ".tsv", ".txt", ".xlsx", ".xlsm")


def save_uploads(files):
    shutil.rmtree(IN_DIR, ignore_errors=True); os.makedirs(IN_DIR)
    paths = []
    def put(name, data):
        name = os.path.basename(name)
        if not name or name.startswith(("~$", ".")) or not name.lower().endswith(EXTS):
            return
        p = os.path.join(IN_DIR, name); i = 2
        while os.path.exists(p):
            b, e = os.path.splitext(name); p = os.path.join(IN_DIR, f"{b} ({i}){e}"); i += 1
        open(p, "wb").write(data); paths.append(p)
    for f in files:
        if f.name.lower().endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(f.getvalue())) as z:
                for info in z.infolist():
                    if not info.is_dir():
                        put(info.filename, z.read(info))
        else:
            put(f.name, f.getvalue())
    return sorted(paths)


def reset(*keys):
    for k in keys:
        SS.pop(k, None)


# ---------------------------------------------------------------- 1. upload
step(1, "Upload files")
files = st.file_uploader("CSV, TSV, TXT, Excel (.xlsx/.xlsm) or a .zip of a folder", type=["csv", "tsv", "txt", "xlsx", "xlsm", "zip"],
                         accept_multiple_files=True, label_visibility="collapsed")
if not files:
    st.info("Upload the files to combine - for example a participant list plus questionnaire, lab and measurement exports. "
            "Every sheet of every workbook is read.")
    footer(); st.stop()
sig = tuple(sorted((f.name, f.size) for f in files))
if SS.get("upload_sig") != sig:
    reset("tables", "skipped", "plan", "result", "exports", "preset")
    SS.paths = save_uploads(files); SS.upload_sig = sig

# ---------------------------------------------------------------- 2. study
step(2, "Study the files")
if "tables" not in SS:
    bar = st.progress(0.0, text="Reading files...")
    tables, skipped = load_tables(SS.paths, progress=lambda i, n, t: bar.progress(min(i / max(n, 1), 1.0), text=f"Reading {t}"))
    bar.empty(); SS.tables, SS.skipped = tables, skipped
tables, skipped = SS.tables, SS.skipped
T = {t.tid: t for t in tables}
c1, c2, c3 = st.columns(3)
c1.metric("Files uploaded", len(SS.paths)); c2.metric("Tables found", len(tables)); c3.metric("Files not readable", len(skipped))
study = pd.DataFrame([dict(table=table_label(t), rows=len(t.data), columns=len(t.data.columns), header_row=t.header_row,
                           empty_rows_dropped=t.empty_rows,
                           likely_id=(t.candidates[0]["column"] if t.candidates else "-"),
                           id_unique=(f"{100 * t.candidates[0]['unique_ratio']:.0f}%" if t.candidates else "-"),
                           numbers_stored_as_text=sum(i.get("text_numbers", 0) for i in t.info.values()),
                           date_columns=", ".join(f"{c} ({i['date_format']}{', ambiguous' if i.get('date_ambiguous') else ''})"
                                                  for c, i in t.info.items() if i.get("date_format")))
                      for t in tables])
st.dataframe(study, hide_index=True, use_container_width=True)
if skipped:
    st.warning("Not read: " + "; ".join(f"{s['file']} - {s['reason']}" for s in skipped))
if not tables:
    st.error("No readable tables were found."); footer(); st.stop()

# ---------------------------------------------------------------- 3. plan
step(3, "Review the merge plan")
opts = base_options(tables)
if not opts:
    st.error("No table has an ID column with unique, complete values. Upload a list with one row per participant / ID.")
    footer(); st.stop()
if "plan" not in SS:
    try:
        SS.plan = make_plan(tables)
    except PlanError as e:
        st.error(str(e)); footer(); st.stop()
plan = SS.plan
b1, b2 = st.columns(2)
base_ids = list(opts)
new_base = b1.selectbox("Base table (one row per ID in the result)", base_ids, index=base_ids.index(plan["base"]),
                        format_func=lambda k: table_label(T[k]))
key_opts = opts[new_base]
new_key = b2.selectbox("ID column in the base table", key_opts,
                       index=key_opts.index(plan["key"]) if plan["key"] in key_opts and new_base == plan["base"] else 0)
if (new_base, new_key) != (plan["base"], plan["key"]):
    SS.plan = make_plan(tables, new_base, new_key); reset("result", "exports", "preset"); st.rerun()

st.caption(f"Base: **{table_label(T[plan['base']])}**, ID column **{plan['key']}** - "
           f"{plan['items'][0]['ids']:,} IDs. Each other table is linked through the column whose values best match these IDs.")
hdr = st.columns([3.2, 1.2, 0.8, 2.2, 2.6])
for h, txt in zip(hdr, ["Table", "Match / grain", "Include", "Linked through", "Rule for several rows per ID"]):
    h.markdown(f"**{txt}**")
changed = False
for it in plan["items"][1:]:
    cols = st.columns([3.2, 1.2, 0.8, 2.2, 2.6])
    cols[0].markdown(f"<div class='ai-row'>{it['label']}<br><small>{it['rows']:,} rows"
                     f"{' &middot; ' + it['note'] if it.get('note') else ''}</small></div>", unsafe_allow_html=True)
    cols[1].markdown(f"<div class='ai-row'>{it['match_pct']:.0f}% match<br><small>{it['grain']}</small></div>", unsafe_allow_html=True)
    inc = cols[2].checkbox("include", value=it["include"], key=f"inc_{it['tid']}", label_visibility="collapsed",
                           disabled=it["role"] == "unlinked")
    if inc != it["include"]:
        it["include"] = inc; changed = True
    if it["link_options"]:
        cur = it["link"] if it["link"] in it["link_options"] else it["link_options"][0]
        link = cols[3].selectbox("link", it["link_options"], index=it["link_options"].index(cur), key=f"link_{it['tid']}",
                                 label_visibility="collapsed")
        if it["role"] == "unlinked" and link == cur and not it.get("link"):
            pass
        elif link != it["link"]:
            update_item(tables, plan, it["tid"], link); changed = True
    else:
        cols[3].markdown("<small>no ID-like column</small>", unsafe_allow_html=True)
    if len(it["rule_options"]) > 1:
        rule = cols[4].selectbox("rule", it["rule_options"], index=it["rule_options"].index(it["rule"]), key=f"rule_{it['tid']}",
                                 label_visibility="collapsed")
        if rule != it["rule"]:
            it["rule"] = rule; changed = True
    else:
        cols[4].markdown(f"<small>{'merge directly' if it['rule'] == 'direct' else 'not linked'}</small>", unsafe_allow_html=True)
if changed:
    reset("result", "exports", "preset"); st.rerun()
drop_empty = st.checkbox("Drop columns that are completely empty", value=True)

# ---------------------------------------------------------------- 4. build
step(4, "Build the analysis file")
if "result" not in SS or SS.get("drop_empty") != drop_empty:
    if st.button("Build", type="primary"):
        bar = st.progress(0.0, text="Merging...")
        try:
            SS.result = execute_plan(tables, plan, os.path.join(WORK, "out"), drop_empty=drop_empty,
                                     progress=lambda i, n, t: bar.progress(min(i / max(n, 1), 1.0), text=f"Merging {t}"))
            SS.drop_empty = drop_empty
            reset("exports", "preset"); st.rerun()
        except PlanError as e:
            bar.empty(); st.error(f"**The build stopped:** {e}")
        except Exception as e:
            bar.empty(); st.error(f"Unexpected error: {type(e).__name__}: {e}")
    footer(); st.stop()

r = SS.result
m = r["merged"]
merged_tables = [p for p in r["profile"].to_dict("records") if str(p["status"]).startswith(("merged", "base"))]
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("IDs (rows)", f"{len(m):,}")
k2.metric("Tables merged", len(merged_tables))
k3.metric("Variables", m.shape[1])
k4.metric("Not merged", len(r["not_merged"]))
k5.metric("Consistency issues", len(r["consistency"]))
st.success(f"One row per **{r['key']}** from {r['base_label']}. Every decision is in the merge log; nothing was dropped silently.")
tabs = st.tabs(["Preview", "File profile", "Merge log", "Consistency check", "Outliers", "Not merged", "Data dictionary"]
               + (["Separate long tables"] if r["separate"] else []))
with tabs[0]:
    st.dataframe(m.head(200), use_container_width=True, height=360); st.caption("First 200 rows.")
with tabs[1]:
    st.dataframe(r["profile"], use_container_width=True, hide_index=True)
with tabs[2]:
    st.dataframe(r["log"], use_container_width=True, hide_index=True)
    st.caption("Columns kept once, renamed because of conflicting values, dropped as copies of the base table, or dropped as empty.")
with tabs[3]:
    st.dataframe(r["consistency"], use_container_width=True, hide_index=True)
    st.caption("Values that differ from the base table in columns that repeat it (e.g. a participant name typed again in another file).")
with tabs[4]:
    st.dataframe(r["outliers"], use_container_width=True, hide_index=True)
    st.caption("Values far from the column median (robust z-score above 7). For checking - nothing is changed.")
with tabs[5]:
    st.dataframe(r["not_merged"], use_container_width=True, hide_index=True)
with tabs[6]:
    st.dataframe(r["dictionary"], use_container_width=True, hide_index=True)
if r["separate"]:
    with tabs[7]:
        for tid, (df, _, label) in r["separate"].items():
            st.markdown(f"**{label}** - {len(df):,} rows"); st.dataframe(df.head(100), use_container_width=True)

# ---------------------------------------------------------------- 5. output
step(5, "Choose your table and download")
owners = [p["tid"] if "tid" in p else None for p in []]
merged_owner_ids = [o for o in dict.fromkeys(r["owner"].values()) if o not in ("_flags", None) and o in T and o != plan["base"]]
o1, o2 = st.columns([2, 1])
with o1:
    chosen = st.multiselect("Tables to include (the base table is always included)", merged_owner_ids, default=merged_owner_ids,
                            format_func=lambda k: table_label(T[k]))
with o2:
    can_helti = helti_applicable(r)
    use_helti = st.toggle("HeLTI analysis variables", value=can_helti, disabled=not can_helti,
                          help="Adds the 73 verified HeLTI / EINSTEIN derived variables (BMI, BP, anaemia, PHQ-9, PSQI, GPAQ...). "
                               "Available when HeLTI columns are present.")
    pii = pii_columns(r)
    deid = st.toggle("De-identify", value=False, disabled=not pii,
                     help="Removes participant identifiers: " + (", ".join(pii) if pii else "none found"))
    all_only = st.toggle("Only IDs present in every selected table", value=False)
key_opts = (tuple(chosen), use_helti, deid, all_only)
if SS.get("out_opts") != key_opts:
    reset("exports"); SS.out_opts = key_opts
preset = None
if use_helti:
    if "preset" not in SS:
        SS.preset = helti_derived(r)
    preset = SS.preset
tbl, kinds, labels, vlabels = build_output(r, tables=set(chosen), preset=preset, deidentify=deid, all_selected_only=all_only)
st.caption(f"Your table: **{len(tbl):,} rows x {tbl.shape[1]} columns**.")
stamp = dt.datetime.now().strftime("%Y%m%d")
if "exports" not in SS:
    if st.button("Prepare downloads", type="primary"):
        with st.spinner("Building CSV, Excel, SPSS and package files..."):
            readme = (f"{B['app_title']} - {B['powered_by']}\n\n"
                      "staging/            one clean CSV per source table (types declared)\n"
                      "merge_syntax.sps    rebuilds the analysis file in SPSS - set the CD line, Run > All;\n"
                      "                    the VERIFICATION table must say PASS on every row\n"
                      + ("helti_analysis_syntax.sps  adds the HeLTI derived variables to analysis_file.sav\n" if preset is not None else ""))
            SS.exports = dict(csv=to_csv(tbl, kinds),
                              xlsx=to_excel(tbl, kinds, labels, r, header_hex=B["PRIMARY"], deidentify=deid,
                                            preset_cols=preset[0].columns if preset is not None else ()),
                              sav=to_sav(tbl, kinds, labels, vlabels, key=r["key"]),
                              zip=to_package(r, preset, readme))
        st.rerun()
else:
    ex = SS.exports
    d1, d2, d3, d4 = st.columns(4)
    d1.download_button("CSV", ex["csv"], f"analysis_file_{stamp}.csv", "text/csv", use_container_width=True)
    d2.download_button("Excel (.xlsx)", ex["xlsx"], f"analysis_file_{stamp}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    d3.download_button("SPSS (.sav)", ex["sav"], f"analysis_file_{stamp}.sav", "application/octet-stream", use_container_width=True)
    d4.download_button("SPSS package (.zip)", ex["zip"], f"analysis_merge_package_{stamp}.zip", "application/zip", use_container_width=True)
    st.caption("Excel includes the data dictionary, file profile, merge log, consistency check, outliers and not-merged sheets.")
footer()
