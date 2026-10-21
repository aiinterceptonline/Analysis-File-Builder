"""Independent Python reference implementation of the HELTI analysis-ready derived variables.
Used only to verify the SPSS syntax value by value."""
import numpy as np, pandas as pd

def nan_if(cond, s):
    s = s.copy(); s[cond] = np.nan; return s

def cat(conds, vals, index):
    out = pd.Series(np.nan, index=index)
    for c, v in zip(conds, vals):
        out[c.fillna(False).astype(bool)] = v
    return out

def flag(cond):
    return cond.fillna(False).astype(bool).astype(float)

def derive(m, name_mismatch_ids):
    d = pd.DataFrame(index=m.index)
    rid = m.record_id
    # ---- cleaned copies
    hb = m.lab_hemoglobin.copy(); hb[(rid == 1054) & (m.lab_hemoglobin == 141)] = 11.4
    d["hb_clean"] = nan_if((hb < 2) | (hb > 20), hb)
    a1 = m.sample_hba1c_m_pc.copy(); a1[(rid == 11066) & (m.sample_hba1c_m_pc == 56)] = 5.0
    d["hba1c_clean"] = nan_if((a1 < 3) | (a1 > 20), a1)
    d["edu_clean"] = nan_if(m.sdc_edu_years > 25, m.sdc_edu_years)
    d["menarche_clean"] = nan_if((m.rep_menstruation_age < 8) | (m.rep_menstruation_age > 20), m.rep_menstruation_age)
    d["income_clean"] = nan_if(m.eru_monthly_income < 500, m.eru_monthly_income)
    items = [f"foodsecurity{i}" for i in range(1, 9)]
    d["food_score_clean"] = nan_if(m[items].isna().all(axis=1), m.foodsecurity_score)
    sedm = m.gpaq16_hours_to_minutes
    d["sedentary_min_day"] = nan_if(sedm > 960, sedm)
    # ---- flags
    d["flag_age"] = flag((m.sdc_age < 18) | (m.sdc_age > 35))
    d["flag_edu"] = flag(m.sdc_edu_years > 20)
    d["flag_menarche"] = flag((m.rep_menstruation_age < 9) | (m.rep_menstruation_age > 18))
    d["flag_bmi_low"] = flag(m.bmi < 15)
    d["flag_height_low"] = flag(m.average_height < 140)
    d["flag_hb"] = flag((m.lab_hemoglobin < 5) | (m.lab_hemoglobin > 18))
    d["flag_hba1c"] = flag((m.sample_hba1c_m_pc < 3.5) | (m.sample_hba1c_m_pc > 15))
    d["flag_sleep_eff"] = flag(m.lsb_sle_c4_habitual_sleep_efficiency > 100)
    d["flag_pulse"] = flag((m.average_pulse < 45) | (m.average_pulse > 130))
    d["flag_income"] = flag(m.eru_monthly_income < 500)
    d["flag_parity"] = flag((m.rep_preg_nb_living >= 5) | (m.birth_still_born >= 4))
    d["flag_name_mismatch"] = flag(rid.isin(name_mismatch_ids))
    fl = [c for c in d.columns if c.startswith("flag_")]
    d["flag_any"] = (d[fl].sum(axis=1) > 0).astype(float)
    # ---- socio-demographic
    age = m.sdc_age
    d["age_grp"] = cat([age < 20, (age >= 20) & (age < 25), (age >= 25) & (age < 30), age >= 30], [1, 2, 3, 4], m.index)
    e = d.edu_clean
    d["edu_cat"] = cat([e == 0, (e >= 1) & (e <= 5), (e >= 6) & (e <= 10), (e >= 11) & (e <= 12), e >= 13], [0, 1, 2, 3, 4], m.index)
    d["income_log"] = np.log(d.income_clean)
    s = m.score_sli_total
    d["sli_cat"] = cat([s <= 14, (s >= 15) & (s <= 24), s >= 25], [1, 2, 3], m.index)
    yes = m[items].eq(1).sum(axis=1).astype(float)
    d["fies_raw"] = nan_if(m[items].isna().any(axis=1), yes)
    f = d.fies_raw
    d["fies_cat"] = cat([f == 0, (f >= 1) & (f <= 3), (f >= 4) & (f <= 6), f >= 7], [0, 1, 2, 3], m.index)
    anyyes = m[items].eq(1).any(axis=1)
    d["food_insecure_any"] = np.where(anyyes, 1.0, np.where(m[items].notna().any(axis=1), 0.0, np.nan))
    # ---- reproductive
    mc = d.menarche_clean
    d["menarche_early"] = np.where(mc.isna(), np.nan, (mc < 12).astype(float))
    lv = m.rep_preg_nb_living
    d["living_cat"] = cat([lv == 0, lv == 1, lv >= 2], [0, 1, 2], m.index)
    d["abortion_any"] = np.where(m.rep_preg_nb_abortion.isna(), np.nan, (m.rep_preg_nb_abortion > 0).astype(float))
    d["stillbirth_any"] = np.where(m.birth_still_born.isna(), np.nan, (m.birth_still_born > 0).astype(float))
    rp = m[["rep_preg_nb_living", "rep_preg_nb_abortion", "birth_still_born"]]
    d["ever_pregnant"] = np.where((rp > 0).any(axis=1), 1.0, np.where(rp.notna().any(axis=1), 0.0, np.nan))
    # ---- anthropometry
    b = m.bmi
    d["bmi_who"] = cat([b < 18.5, (b >= 18.5) & (b < 25), (b >= 25) & (b < 30), b >= 30], [1, 2, 3, 4], m.index)
    d["bmi_asian"] = cat([b < 18.5, (b >= 18.5) & (b < 23), (b >= 23) & (b < 27.5), b >= 27.5], [1, 2, 3, 4], m.index)
    d["underweight"] = np.where(b.isna(), np.nan, (b < 18.5).astype(float))
    h = m.average_height
    d["short_stature"] = np.where(h.isna(), np.nan, (h < 145).astype(float))
    # ---- blood pressure
    sb, db = m.pm_sbp, m.pm_dbp
    htn = pd.Series(np.nan, index=m.index)
    htn[sb.notna() & db.notna()] = 0
    htn[(sb >= 140) | (db >= 90)] = 1
    d["htn"] = htn
    both = sb.notna() & db.notna()
    bc = pd.Series(np.nan, index=m.index)
    bc[both & (sb < 120) & (db < 80)] = 1
    bc[both & (sb >= 120) & (sb < 130) & (db < 80)] = 2
    bc[both & (((sb >= 130) & (sb < 140)) | ((db >= 80) & (db < 90)))] = 3
    bc[both & ((sb >= 140) | (db >= 90))] = 4
    d["bp_cat"] = bc
    # ---- labs
    hc = d.hb_clean
    d["anaemia"] = np.where(hc.isna(), np.nan, (hc < 12).astype(float))
    d["anaemia_sev"] = cat([hc >= 12, (hc >= 11) & (hc < 12), (hc >= 8) & (hc < 11), hc < 8], [0, 1, 2, 3], m.index)
    ac = d.hba1c_clean
    d["glyc_cat"] = cat([ac < 5.7, (ac >= 5.7) & (ac < 6.5), ac >= 6.5], [1, 2, 3], m.index)
    d["rbs_high"] = np.where(m.lab_rbs.isna(), np.nan, (m.lab_rbs >= 200).astype(float))
    dm = pd.Series(np.nan, index=m.index)
    dm[ac.notna() | m.lab_rbs.notna()] = 0
    dm[(ac >= 6.5) | (m.lab_rbs >= 200)] = 1
    d["diabetes_any"] = dm
    # ---- psychological
    p = m.psy_dep_phq9_score
    d["phq9_cat"] = cat([p <= 4, (p >= 5) & (p <= 9), (p >= 10) & (p <= 14), (p >= 15) & (p <= 19), p >= 20], [0, 1, 2, 3, 4], m.index)
    d["depression"] = np.where(p.isna(), np.nan, (p >= 10).astype(float))
    i9 = m.psy_dep_phq9_9
    d["phq9_item9_pos"] = np.where(i9.isna(), np.nan, (i9 > 0).astype(float))
    # ---- sleep (PSQI, standard scoring, REDCap item mapping)
    q2, q4 = m.lsb_sle_psqi2, m.lsb_sle_psqi4
    c1 = m.lsb_sle_psqi9
    lat = cat([q2 <= 15, (q2 > 15) & (q2 <= 30), (q2 > 30) & (q2 <= 60), q2 > 60], [0, 1, 2, 3], m.index)
    s2 = lat + m.lsb_sle_psqi5a
    c2 = cat([s2 == 0, (s2 >= 1) & (s2 <= 2), (s2 >= 3) & (s2 <= 4), s2 >= 5], [0, 1, 2, 3], m.index)
    c3 = cat([q4 > 7, (q4 >= 6) & (q4 <= 7), (q4 >= 5) & (q4 < 6), q4 < 5], [0, 1, 2, 3], m.index)
    hib = m.lsb_sle_c4_hours_spent_calculation
    eff = (q4 / hib * 100).where(hib > 0)
    c4 = cat([eff >= 85, (eff >= 75) & (eff < 85), (eff >= 65) & (eff < 75), eff < 65], [0, 1, 2, 3], m.index)
    dist = m[[f"lsb_sle_psqi5{x}" for x in "bcdefghij"]].sum(axis=1, min_count=1)
    c5 = cat([dist == 0, (dist >= 1) & (dist <= 9), (dist >= 10) & (dist <= 18), dist >= 19], [0, 1, 2, 3], m.index)
    c6 = m.lsb_sle_psqi6
    s7 = m.lsb_sle_psqi7 + m.lsb_sle_psqi8
    c7 = cat([s7 == 0, (s7 >= 1) & (s7 <= 2), (s7 >= 3) & (s7 <= 4), s7 >= 5], [0, 1, 2, 3], m.index)
    for k, v in enumerate([c1, c2, c3, c4, c5, c6, c7], 1):
        d[f"psqi_c{k}_std"] = v
    comps = d[[f"psqi_c{k}_std" for k in range(1, 8)]]
    d["psqi_global_std"] = comps.sum(axis=1).where(comps.notna().all(axis=1))
    d["poor_sleep"] = np.where(d.psqi_global_std.isna(), np.nan, (d.psqi_global_std > 5).astype(float))
    d["sleep_eff_pct"] = np.minimum(eff, 100)
    d["sleep_dur_cat"] = cat([q4 < 7, (q4 >= 7) & (q4 <= 9), q4 > 9], [1, 2, 3], m.index)
    # ---- physical activity (GPAQ)
    def dom(yesq, daysq, minq):
        Y, D, M = m[yesq], m[daysq], m[minq]
        out = pd.Series(np.nan, index=m.index)
        out[Y == 2] = 0
        ok = (Y == 1) & D.between(1, 7) & M.notna() & (M <= 960)
        out[ok] = (D * M)[ok]
        return out
    d["pa_vig_work_min"] = dom("lsb_pa_gpaq1", "lsb_pa_gpaq2", "gpaq3_hours_to_minutes")
    d["pa_mod_work_min"] = dom("lsb_pa_gpaq4", "lsb_pa_gpaq5", "gpaq6_hours_to_minutes")
    d["pa_travel_min"] = dom("lsb_pa_gpaq7", "lsb_pa_gpaq8", "gpaq9a_hours_to_minutes")
    d["pa_vig_rec_min"] = dom("lsb_pa_gpaq10", "lsb_pa_gpaq11", "gpaq12_hours_to_minutes")
    d["pa_mod_rec_min"] = dom("lsb_pa_gpaq13", "lsb_pa_gpaq14", "gpaq15_hours_to_minutes")
    met = 8 * d.pa_vig_work_min + 4 * d.pa_mod_work_min + 4 * d.pa_travel_min + 8 * d.pa_vig_rec_min + 4 * d.pa_mod_rec_min
    d["pa_met_total"] = met
    d["pa_meets_who"] = np.where(met.isna(), np.nan, (met >= 600).astype(float))
    # ---- completeness
    d["has_anthro"] = (m.average_height.notna() & m.average_weight.notna()).astype(float)
    d["has_bp"] = (m.pm_sbp.notna() & m.pm_dbp.notna()).astype(float)
    d["has_hb"] = d.hb_clean.notna().astype(float)
    d["has_hba1c"] = d.hba1c_clean.notna().astype(float)
    d["has_phq9"] = m.psy_dep_phq9_score.notna().astype(float)
    d["has_gse"] = m.psy_efficacy_gse_score.notna().astype(float)
    d["has_psqi"] = d.psqi_global_std.notna().astype(float)
    d["has_repro"] = m.rep_menstruation_age.notna().astype(float)
    core = ["has_anthro", "has_bp", "has_hb", "has_phq9", "has_gse", "has_psqi", "has_repro"]
    d["complete_core"] = (d[core].sum(axis=1) == len(core)).astype(float)
    return d
