"""Output table options + CSV / Excel / SPSS / package builders for the generic engine."""
import datetime as dt, io, json, os, re, tempfile, zipfile
import numpy as np
import pandas as pd

from .generic import PII_NAME
from .derive import derive
from .analysis_syntax import make_analysis_syntax

HERE = os.path.dirname(os.path.abspath(__file__))
DMETA = json.load(open(os.path.join(HERE, "derived_labels.json")))
DERIVE_INPUTS = [
    "record_id", "lab_hemoglobin", "sample_hba1c_m_pc", "sdc_edu_years", "rep_menstruation_age", "eru_monthly_income",
    "foodsecurity_score", "gpaq16_hours_to_minutes", "sdc_age", "bmi", "average_height", "average_weight",
    "lsb_sle_c4_habitual_sleep_efficiency", "average_pulse", "rep_preg_nb_living", "birth_still_born",
    "rep_preg_nb_abortion", "score_sli_total", "pm_sbp", "pm_dbp", "lab_rbs", "psy_dep_phq9_score", "psy_dep_phq9_9",
    "psy_efficacy_gse_score", "lsb_sle_psqi2", "lsb_sle_psqi4", "lsb_sle_psqi5a", "lsb_sle_psqi6", "lsb_sle_psqi7",
    "lsb_sle_psqi8", "lsb_sle_psqi9", "lsb_sle_c4_hours_spent_calculation",
] + [f"foodsecurity{i}" for i in range(1, 9)] + [f"lsb_sle_psqi5{x}" for x in "bcdefghij"] + \
    [f"lsb_pa_gpaq{i}" for i in (1, 2, 4, 5, 7, 8, 10, 11, 13, 14)] + \
    ["gpaq3_hours_to_minutes", "gpaq6_hours_to_minutes", "gpaq9a_hours_to_minutes", "gpaq12_hours_to_minutes", "gpaq15_hours_to_minutes"]
HELTI_SIGNATURE = ["lab_hemoglobin", "psy_dep_phq9_score", "lsb_sle_psqi4", "average_height", "average_weight",
                   "pm_sbp", "rep_menstruation_age", "lsb_pa_gpaq1", "foodsecurity_score", "sdc_age"]


# ------------------------------------------------------------------------------------------- HeLTI preset
def helti_applicable(result):
    m = result["merged"]
    return result["key"] == "record_id" and sum(c in m.columns for c in HELTI_SIGNATURE) >= 5


def helti_derived(result):
    """The 73 HeLTI analysis variables (same definitions as HELTI_analysis_syntax.sps)."""
    m = result["merged"].copy()
    if "bmi" not in m and {"average_weight", "average_height"} <= set(m.columns):
        m["bmi"] = (m["average_weight"] / (m["average_height"] / 100) ** 2).round(1)
    for c in DERIVE_INPUTS:
        if c not in m:
            m[c] = np.nan
    q = result["consistency"]
    ids = sorted({int(i) for i in q["id"]}) if len(q) else []
    d = derive(m, ids)
    order = [c for c in DMETA["order"] if c in d.columns]
    keep = [c for c in order if d[c].notna().any() and not (c.startswith("has_") and d[c].sum() == 0)]
    d = d[keep]
    if "bmi" not in result["merged"] and "bmi" in m:
        d.insert(0, "bmi", m["bmi"].values)
    return d, ids


# ------------------------------------------------------------------------------------------- table builder
STAFF = re.compile(r"interviewer|user|staff|employment|occupation|enumerator|data_?entry|collector|doctor|village|school", re.I)


def pii_columns(result):
    """Text columns that identify a participant or family member (names, phone, address...). Staff names are kept."""
    return [c for c in result["merged"].columns if result["kinds"].get(c) == "text" and PII_NAME.search(c)
            and not STAFF.search(c) and c != result["key"]]


def build_output(result, tables=None, preset=None, deidentify=False, all_selected_only=False):
    m = result["merged"]; owner = result["owner"]; key = result["key"]
    base_owner = owner.get(key)
    cols = []
    for c in m.columns:
        o = owner.get(c)
        if c == key or tables is None or o in (base_owner, "_flags") or o in tables:
            if o == "_flags" and tables is not None and c.startswith("has_") and c[4:] not in tables and c[4:] != base_owner and c != "n_datasets":
                continue
            cols.append(c)
    if deidentify:
        pii = set(pii_columns(result)); cols = [c for c in cols if c not in pii]
    t = m[cols].copy()
    kinds = dict(result["kinds"]); labels = {c: result["labels"].get(c) for c in cols}; vlabels = {}
    if preset is not None:
        d, _ = preset
        t = pd.concat([t.reset_index(drop=True), d.reset_index(drop=True)], axis=1)
        for c in d.columns:
            kinds[c] = "num"; labels[c] = DMETA["labels"].get(c) or ("BMI (kg/m2) = weight / (height/100)^2" if c == "bmi" else c)
            if c in DMETA["value_labels"]:
                vlabels[c] = {float(k): v for k, v in DMETA["value_labels"][c].items()}
        for k_, v_ in DMETA["psqi_item_labels"].items():
            if k_ in labels: labels[k_] = v_
    if all_selected_only:
        flags = [c for c in t.columns if c.startswith("has_")]
        if flags:
            t = t[(t[flags] == 1).all(axis=1)].reset_index(drop=True)
    return t, kinds, labels, vlabels


# ------------------------------------------------------------------------------------------- writers
def _flat(t, kinds):
    x = t.copy()
    for c in x.columns:
        k = kinds.get(c)
        if k == "date":
            x[c] = [v if isinstance(v, dt.date) and not pd.isna(v) else None for v in x[c]]
        elif k == "time":
            x[c] = [v.strftime("%H:%M") if isinstance(v, dt.time) else None for v in x[c]]
    return x


def to_csv(t, kinds):
    x = _flat(t, kinds)
    for c in x.columns:
        if kinds.get(c) == "date":
            x[c] = [v.isoformat() if isinstance(v, dt.date) else "" for v in x[c]]
    return x.to_csv(index=False).encode("utf-8-sig")


def to_excel(t, kinds, labels, result, header_hex="0F2A4A", deidentify=False, preset_cols=()):
    x = _flat(t, kinds)
    for c in x.columns:
        if kinds.get(c) == "date":
            x[c] = [dt.datetime.combine(v, dt.time()) if isinstance(v, dt.date) else None for v in x[c]]
    d = result["dictionary"].set_index("variable")
    dictionary = pd.DataFrame([dict(variable=c, label=labels.get(c) or "", type=kinds.get(c),
                                    source_file=("derived (HeLTI preset)" if c in preset_cols else d["source_file"].get(c, "")),
                                    original_column=("" if c in preset_cols else d["original_column"].get(c, "")),
                                    n_filled=int(t[c].notna().sum())) for c in t.columns])
    cons = result["consistency"].copy()
    if deidentify and len(cons):
        cons = cons.drop(columns=[c for c in ("base_value", "file_value") if c in cons.columns])
    sheets = [("Data", x), ("Data_dictionary", dictionary), ("File_profile", result["profile"]), ("Merge_log", result["log"]),
              ("Consistency_check", cons), ("Outliers", result["outliers"]), ("Files_not_merged", result["not_merged"])]
    for tid, (df, k2, label) in result["separate"].items():
        sheets.append((("Long_" + tid)[:31], _flat(df, k2)))
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter", datetime_format="dd-mm-yyyy", date_format="dd-mm-yyyy") as w:
        head = w.book.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#" + header_hex.lstrip("#")})
        for name, df in sheets:
            if df is None or (len(df.columns) == 0):
                df = pd.DataFrame({"info": ["nothing to report"]})
            df.to_excel(w, sheet_name=name, index=False)
            ws = w.sheets[name]
            for j, col in enumerate(df.columns):
                ws.write(0, j, col, head)
                sample = max([len(str(v)) for v in df[col].head(150) if v is not None and v == v] or [0])
                ws.set_column(j, j, min(max(10, len(str(col)), sample) + 2, 48))
            ws.freeze_panes(1, 1 if name == "Data" else 0)
    return buf.getvalue()


def to_sav(t, kinds, labels, vlabels, key=None):
    import pyreadstat
    sp = t.copy(); vfmt = {}; measure = {key: "nominal"} if key else {}
    for c in sp.columns:
        k = kinds.get(c)
        if k == "date":
            sp[c] = pd.to_datetime(sp[c]).astype("datetime64[ns]"); vfmt[c] = "DATE11"
        elif k == "datetime":
            sp[c] = pd.to_datetime(sp[c]).astype("datetime64[ns]"); vfmt[c] = "DATETIME20"
        elif k == "time":
            sp[c] = [v.hour * 3600 + v.minute * 60 + v.second if isinstance(v, dt.time) else np.nan for v in sp[c]]; vfmt[c] = "TIME5"
        elif k in ("text", "empty") and sp[c].dtype != object:
            sp[c] = sp[c].astype(object)
        if c in vlabels:
            measure[c] = "nominal" if set(vlabels[c]) <= {0.0, 1.0} else "ordinal"
    with tempfile.TemporaryDirectory() as d:
        p = os.path.join(d, "out.sav")
        pyreadstat.write_sav(sp, p, column_labels={c: (labels.get(c) or None) for c in sp.columns}, variable_format=vfmt,
                             variable_value_labels=vlabels, variable_measure=measure)
        return open(p, "rb").read()


def to_package(result, preset=None, readme=""):
    out = result["out_dir"]; buf = io.BytesIO(); base = "analysis_merge_package"
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(os.listdir(os.path.join(out, "staging"))):
            z.write(os.path.join(out, "staging", f), f"{base}/staging/{f}")
        z.write(os.path.join(out, "merge_syntax.sps"), f"{base}/merge_syntax.sps")
        if preset is not None:
            d, ids = preset
            s = make_analysis_syntax(d.drop(columns=["bmi"], errors="ignore"), ids, cd_folder="C:\\analysis\\merge_package")
            s = s.replace("GET FILE='HELTI_merged_baseline.sav'.", "GET FILE='analysis_file.sav'.\n"
                          "* BMI from height and weight (the generic merge does not add it).\n"
                          "COMPUTE bmi = RND(10 * average_weight / (average_height / 100) ** 2) / 10.\nEXECUTE.")
            s = s.replace("SAVE OUTFILE='HELTI_analysis_baseline.sav'", "SAVE OUTFILE='analysis_file_with_helti_variables.sav'")
            s = s.replace("SAVE TRANSLATE OUTFILE='HELTI_analysis_baseline.xlsx'", "SAVE TRANSLATE OUTFILE='analysis_file_with_helti_variables.xlsx'")
            z.writestr(f"{base}/helti_analysis_syntax.sps", s)
        if readme:
            z.writestr(f"{base}/README.txt", readme)
    return buf.getvalue()
