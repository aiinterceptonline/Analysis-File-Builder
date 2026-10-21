"""SPSS syntax for the analysis-ready layer. The verification counts and name-check IDs are filled in
from the current run, so the syntax always matches the data it was generated with."""

TEMPLATE = r"""* Encoding: UTF-8.
* =====================================================================.
* HELTI / EINSTEIN baseline - ANALYSIS-READY FILE.
* Version 1 (20 Oct 2026).
*
* INPUT : HELTI_merged_baseline.sav  (5,643 women x 189 variables; merged file).
*         The SPSS-produced HELTI_merged_baseline_spss.sav works too - change the GET FILE line.
* OUTPUT: HELTI_analysis_baseline.sav and HELTI_analysis_baseline.xlsx
*
* RULES
*  - Raw variables are NEVER changed. Corrections go into new *_clean variables.
*  - flag_* variables mark values to check (1 = check); nothing is deleted because of a flag.
*  - Derived variables are system-missing when their inputs are missing.
*  - Cut-offs are standard international definitions (WHO etc.). Align with the HeLTI
*    harmonised protocol when it is available - each definition is one block below.
*
* NOTE ON PSQI ITEM NAMES (REDCap numbering differs from the standard PSQI questionnaire):
*   lsb_sle_psqi9 = Q6 subjective sleep quality      lsb_sle_psqi6 = Q7 sleep medication
*   lsb_sle_psqi7 = Q8 trouble staying awake          lsb_sle_psqi8 = Q9 enthusiasm
*   (verified: with this mapping the source components agree with standard scoring).
*
* HOW TO USE: change the CD line, Run > All, check the VERIFICATION table (all PASS).
* =====================================================================.

CD 'C:\HELTI\HELTI_merge_v5'.
GET FILE='HELTI_merged_baseline.sav'.
DATASET NAME analysis.

* =====================================================================.
* SECTION 1. CLEANED COPIES (raw variable kept unchanged next to each).
* =====================================================================.
* Hb: apply the documented lab-report correction (001054: 141 -> 11.4); impossible values (<2 or >20) -> missing.
COMPUTE hb_clean = lab_hemoglobin.
IF (record_id = 1054 AND lab_hemoglobin = 141) hb_clean = 11.4.
IF (hb_clean < 2 OR hb_clean > 20) hb_clean = $SYSMIS.
* HbA1c: documented correction (011066: 56 -> 5.0); impossible values (<3 or >20) -> missing.
COMPUTE hba1c_clean = sample_hba1c_m_pc.
IF (record_id = 11066 AND sample_hba1c_m_pc = 56) hba1c_clean = 5.0.
IF (hba1c_clean < 3 OR hba1c_clean > 20) hba1c_clean = $SYSMIS.
* Education years: values > 25 are calendar years / typing errors (e.g. 2011, 10000) -> missing.
COMPUTE edu_clean = sdc_edu_years.
IF (edu_clean > 25) edu_clean = $SYSMIS.
* Age at menarche: < 8 or > 20 years -> missing.
COMPUTE menarche_clean = rep_menstruation_age.
IF (menarche_clean < 8 OR menarche_clean > 20) menarche_clean = $SYSMIS.
* Monthly household income: < 500 rupees -> missing.
COMPUTE income_clean = eru_monthly_income.
IF (income_clean < 500) income_clean = $SYSMIS.
* Food insecurity score: missing when all 8 items are blank (REDCap stored 0).
COMPUTE food_score_clean = foodsecurity_score.
IF (NMISS(foodsecurity1, foodsecurity2, foodsecurity3, foodsecurity4, foodsecurity5, foodsecurity6, foodsecurity7, foodsecurity8) = 8) food_score_clean = $SYSMIS.
* Sedentary time (GPAQ Q16) in minutes/day: > 960 (16 h) -> missing (WHO GPAQ cleaning rule).
COMPUTE sedentary_min_day = gpaq16_hours_to_minutes.
IF (sedentary_min_day > 960) sedentary_min_day = $SYSMIS.
EXECUTE.

* =====================================================================.
* SECTION 2. DATA-QUALITY FLAGS (1 = please check; 0 = OK or not measured).
* =====================================================================.
COMPUTE flag_age = 0.
IF (sdc_age < 18 OR sdc_age > 35) flag_age = 1.
COMPUTE flag_edu = 0.
IF (sdc_edu_years > 20) flag_edu = 1.
COMPUTE flag_menarche = 0.
IF (rep_menstruation_age < 9 OR rep_menstruation_age > 18) flag_menarche = 1.
COMPUTE flag_bmi_low = 0.
IF (bmi < 15) flag_bmi_low = 1.
COMPUTE flag_height_low = 0.
IF (average_height < 140) flag_height_low = 1.
COMPUTE flag_hb = 0.
IF (lab_hemoglobin < 5 OR lab_hemoglobin > 18) flag_hb = 1.
COMPUTE flag_hba1c = 0.
IF (sample_hba1c_m_pc < 3.5 OR sample_hba1c_m_pc > 15) flag_hba1c = 1.
COMPUTE flag_sleep_eff = 0.
IF (lsb_sle_c4_habitual_sleep_efficiency > 100) flag_sleep_eff = 1.
COMPUTE flag_pulse = 0.
IF (average_pulse < 45 OR average_pulse > 130) flag_pulse = 1.
COMPUTE flag_income = 0.
IF (eru_monthly_income < 500) flag_income = 1.
COMPUTE flag_parity = 0.
IF (rep_preg_nb_living >= 5 OR birth_still_born >= 4) flag_parity = 1.
* Participants whose name / husband name in a data file does not match the master (QC_name_mismatch sheet).
COMPUTE flag_name_mismatch = 0.
DO REPEAT id = @IDS@.
IF (record_id = id) flag_name_mismatch = 1.
END REPEAT.
COMPUTE flag_any = 0.
IF (SUM(flag_age, flag_edu, flag_menarche, flag_bmi_low, flag_height_low, flag_hb, flag_hba1c,
        flag_sleep_eff, flag_pulse, flag_income, flag_parity, flag_name_mismatch) > 0) flag_any = 1.
EXECUTE.

* =====================================================================.
* SECTION 3. SOCIO-DEMOGRAPHIC.
* =====================================================================.
COMPUTE age_grp = $SYSMIS.
IF (sdc_age < 20) age_grp = 1.
IF (sdc_age >= 20 AND sdc_age < 25) age_grp = 2.
IF (sdc_age >= 25 AND sdc_age < 30) age_grp = 3.
IF (sdc_age >= 30) age_grp = 4.
COMPUTE edu_cat = $SYSMIS.
IF (edu_clean = 0) edu_cat = 0.
IF (edu_clean >= 1 AND edu_clean <= 5) edu_cat = 1.
IF (edu_clean >= 6 AND edu_clean <= 10) edu_cat = 2.
IF (edu_clean >= 11 AND edu_clean <= 12) edu_cat = 3.
IF (edu_clean >= 13) edu_cat = 4.
COMPUTE income_log = LN(income_clean).
* Standard of Living Index - NFHS-2 cut-offs (low 0-14, medium 15-24, high 25-67).
COMPUTE sli_cat = $SYSMIS.
IF (score_sli_total <= 14) sli_cat = 1.
IF (score_sli_total >= 15 AND score_sli_total <= 24) sli_cat = 2.
IF (score_sli_total >= 25) sli_cat = 3.
* Food insecurity: raw score = number of YES answers to the 8 items (0-8), only when all 8 answered.
COUNT fies_raw = foodsecurity1 foodsecurity2 foodsecurity3 foodsecurity4 foodsecurity5 foodsecurity6 foodsecurity7 foodsecurity8 (1).
IF (NMISS(foodsecurity1, foodsecurity2, foodsecurity3, foodsecurity4, foodsecurity5, foodsecurity6, foodsecurity7, foodsecurity8) > 0) fies_raw = $SYSMIS.
* Provisional categories (FIES raw-score convention) - confirm the instrument with the study team.
COMPUTE fies_cat = $SYSMIS.
IF (fies_raw = 0) fies_cat = 0.
IF (fies_raw >= 1 AND fies_raw <= 3) fies_cat = 1.
IF (fies_raw >= 4 AND fies_raw <= 6) fies_cat = 2.
IF (fies_raw >= 7) fies_cat = 3.
COMPUTE food_insecure_any = $SYSMIS.
IF (NVALID(foodsecurity1, foodsecurity2, foodsecurity3, foodsecurity4, foodsecurity5, foodsecurity6, foodsecurity7, foodsecurity8) > 0) food_insecure_any = 0.
IF (ANY(1, foodsecurity1, foodsecurity2, foodsecurity3, foodsecurity4, foodsecurity5, foodsecurity6, foodsecurity7, foodsecurity8)) food_insecure_any = 1.
EXECUTE.

* =====================================================================.
* SECTION 4. REPRODUCTIVE HISTORY.
* =====================================================================.
COMPUTE menarche_early = $SYSMIS.
IF (NOT MISSING(menarche_clean)) menarche_early = (menarche_clean < 12).
COMPUTE living_cat = $SYSMIS.
IF (rep_preg_nb_living = 0) living_cat = 0.
IF (rep_preg_nb_living = 1) living_cat = 1.
IF (rep_preg_nb_living >= 2) living_cat = 2.
COMPUTE abortion_any = $SYSMIS.
IF (NOT MISSING(rep_preg_nb_abortion)) abortion_any = (rep_preg_nb_abortion > 0).
COMPUTE stillbirth_any = $SYSMIS.
IF (NOT MISSING(birth_still_born)) stillbirth_any = (birth_still_born > 0).
COMPUTE ever_pregnant = $SYSMIS.
IF (NVALID(rep_preg_nb_living, rep_preg_nb_abortion, birth_still_born) > 0) ever_pregnant = 0.
IF (rep_preg_nb_living > 0 OR rep_preg_nb_abortion > 0 OR birth_still_born > 0) ever_pregnant = 1.
EXECUTE.

* =====================================================================.
* SECTION 5. ANTHROPOMETRY.
* =====================================================================.
* WHO BMI categories.
COMPUTE bmi_who = $SYSMIS.
IF (bmi < 18.5) bmi_who = 1.
IF (bmi >= 18.5 AND bmi < 25) bmi_who = 2.
IF (bmi >= 25 AND bmi < 30) bmi_who = 3.
IF (bmi >= 30) bmi_who = 4.
* WHO Asian cut-offs (2004 expert consultation: 23 and 27.5).
COMPUTE bmi_asian = $SYSMIS.
IF (bmi < 18.5) bmi_asian = 1.
IF (bmi >= 18.5 AND bmi < 23) bmi_asian = 2.
IF (bmi >= 23 AND bmi < 27.5) bmi_asian = 3.
IF (bmi >= 27.5) bmi_asian = 4.
COMPUTE underweight = $SYSMIS.
IF (NOT MISSING(bmi)) underweight = (bmi < 18.5).
COMPUTE short_stature = $SYSMIS.
IF (NOT MISSING(average_height)) short_stature = (average_height < 145).
EXECUTE.

* =====================================================================.
* SECTION 6. BLOOD PRESSURE (pm_sbp / pm_dbp = mean of readings 2 and 3).
* =====================================================================.
* Hypertension: SBP >= 140 or DBP >= 90.
COMPUTE htn = $SYSMIS.
IF (NVALID(pm_sbp, pm_dbp) = 2) htn = 0.
IF (pm_sbp >= 140 OR pm_dbp >= 90) htn = 1.
* ACC/AHA 2017 categories (needs both readings).
COMPUTE bp_cat = $SYSMIS.
DO IF (NVALID(pm_sbp, pm_dbp) = 2).
+  IF (pm_sbp < 120 AND pm_dbp < 80) bp_cat = 1.
+  IF (pm_sbp >= 120 AND pm_sbp < 130 AND pm_dbp < 80) bp_cat = 2.
+  IF ((pm_sbp >= 130 AND pm_sbp < 140) OR (pm_dbp >= 80 AND pm_dbp < 90)) bp_cat = 3.
+  IF (pm_sbp >= 140 OR pm_dbp >= 90) bp_cat = 4.
END IF.
EXECUTE.

* =====================================================================.
* SECTION 7. LABORATORY.
* =====================================================================.
* Anaemia - WHO, non-pregnant women: Hb < 12 g/dL; mild 11-11.9, moderate 8-10.9, severe < 8.
COMPUTE anaemia = $SYSMIS.
IF (NOT MISSING(hb_clean)) anaemia = (hb_clean < 12).
COMPUTE anaemia_sev = $SYSMIS.
IF (hb_clean >= 12) anaemia_sev = 0.
IF (hb_clean >= 11 AND hb_clean < 12) anaemia_sev = 1.
IF (hb_clean >= 8 AND hb_clean < 11) anaemia_sev = 2.
IF (hb_clean < 8) anaemia_sev = 3.
* Glycaemic status from HbA1c (ADA): normal < 5.7, prediabetes 5.7-6.4, diabetes >= 6.5.
COMPUTE glyc_cat = $SYSMIS.
IF (hba1c_clean < 5.7) glyc_cat = 1.
IF (hba1c_clean >= 5.7 AND hba1c_clean < 6.5) glyc_cat = 2.
IF (hba1c_clean >= 6.5) glyc_cat = 3.
COMPUTE rbs_high = $SYSMIS.
IF (NOT MISSING(lab_rbs)) rbs_high = (lab_rbs >= 200).
* Diabetes by either test (HbA1c >= 6.5 or RBS >= 200).
COMPUTE diabetes_any = $SYSMIS.
IF (NVALID(hba1c_clean, lab_rbs) > 0) diabetes_any = 0.
IF (hba1c_clean >= 6.5 OR lab_rbs >= 200) diabetes_any = 1.
EXECUTE.

* =====================================================================.
* SECTION 8. PSYCHOLOGICAL.
* =====================================================================.
* PHQ-9 severity (Kroenke 2001): 0-4 none, 5-9 mild, 10-14 moderate, 15-19 moderately severe, 20-27 severe.
COMPUTE phq9_cat = $SYSMIS.
IF (psy_dep_phq9_score <= 4) phq9_cat = 0.
IF (psy_dep_phq9_score >= 5 AND psy_dep_phq9_score <= 9) phq9_cat = 1.
IF (psy_dep_phq9_score >= 10 AND psy_dep_phq9_score <= 14) phq9_cat = 2.
IF (psy_dep_phq9_score >= 15 AND psy_dep_phq9_score <= 19) phq9_cat = 3.
IF (psy_dep_phq9_score >= 20) phq9_cat = 4.
COMPUTE depression = $SYSMIS.
IF (NOT MISSING(psy_dep_phq9_score)) depression = (psy_dep_phq9_score >= 10).
* PHQ-9 item 9 (thoughts of self-harm) - SAFETY flag.
COMPUTE phq9_item9_pos = $SYSMIS.
IF (NOT MISSING(psy_dep_phq9_9)) phq9_item9_pos = (psy_dep_phq9_9 > 0).
EXECUTE.

* =====================================================================.
* SECTION 9. SLEEP - PSQI rescored with the standard algorithm (Buysse 1989).
* Source scores had small errors (missing sleep hours scored 3; efficiency of exactly 85% scored 1).
* =====================================================================.
* C1 subjective sleep quality.
COMPUTE psqi_c1_std = lsb_sle_psqi9.
* C2 latency: Q2 minutes scored 0-3, plus Q5a.
COMPUTE #lat = $SYSMIS.
IF (lsb_sle_psqi2 <= 15) #lat = 0.
IF (lsb_sle_psqi2 > 15 AND lsb_sle_psqi2 <= 30) #lat = 1.
IF (lsb_sle_psqi2 > 30 AND lsb_sle_psqi2 <= 60) #lat = 2.
IF (lsb_sle_psqi2 > 60) #lat = 3.
COMPUTE #s2 = #lat + lsb_sle_psqi5a.
COMPUTE psqi_c2_std = $SYSMIS.
IF (#s2 = 0) psqi_c2_std = 0.
IF (#s2 >= 1 AND #s2 <= 2) psqi_c2_std = 1.
IF (#s2 >= 3 AND #s2 <= 4) psqi_c2_std = 2.
IF (#s2 >= 5) psqi_c2_std = 3.
* C3 duration (Q4 hours of actual sleep).
COMPUTE psqi_c3_std = $SYSMIS.
IF (lsb_sle_psqi4 > 7) psqi_c3_std = 0.
IF (lsb_sle_psqi4 >= 6 AND lsb_sle_psqi4 <= 7) psqi_c3_std = 1.
IF (lsb_sle_psqi4 >= 5 AND lsb_sle_psqi4 < 6) psqi_c3_std = 2.
IF (lsb_sle_psqi4 < 5) psqi_c3_std = 3.
* C4 habitual sleep efficiency = hours slept / hours in bed x 100.
COMPUTE #eff = $SYSMIS.
IF (lsb_sle_c4_hours_spent_calculation > 0) #eff = lsb_sle_psqi4 / lsb_sle_c4_hours_spent_calculation * 100.
COMPUTE psqi_c4_std = $SYSMIS.
IF (#eff >= 85) psqi_c4_std = 0.
IF (#eff >= 75 AND #eff < 85) psqi_c4_std = 1.
IF (#eff >= 65 AND #eff < 75) psqi_c4_std = 2.
IF (#eff < 65) psqi_c4_std = 3.
* Efficiency for analysis, capped at 100% (reported sleep sometimes exceeds time in bed).
* MIN ignores missing values, so the cap is applied only when efficiency is known.
COMPUTE sleep_eff_pct = $SYSMIS.
IF (NOT MISSING(#eff)) sleep_eff_pct = MIN(#eff, 100).
* C5 disturbances: sum of Q5b-Q5j.
COMPUTE #dist = SUM(lsb_sle_psqi5b, lsb_sle_psqi5c, lsb_sle_psqi5d, lsb_sle_psqi5e, lsb_sle_psqi5f,
                    lsb_sle_psqi5g, lsb_sle_psqi5h, lsb_sle_psqi5i, lsb_sle_psqi5j).
COMPUTE psqi_c5_std = $SYSMIS.
IF (#dist = 0) psqi_c5_std = 0.
IF (#dist >= 1 AND #dist <= 9) psqi_c5_std = 1.
IF (#dist >= 10 AND #dist <= 18) psqi_c5_std = 2.
IF (#dist >= 19) psqi_c5_std = 3.
* C6 use of sleep medication.
COMPUTE psqi_c6_std = lsb_sle_psqi6.
* C7 daytime dysfunction: trouble staying awake + enthusiasm.
COMPUTE #s7 = lsb_sle_psqi7 + lsb_sle_psqi8.
COMPUTE psqi_c7_std = $SYSMIS.
IF (#s7 = 0) psqi_c7_std = 0.
IF (#s7 >= 1 AND #s7 <= 2) psqi_c7_std = 1.
IF (#s7 >= 3 AND #s7 <= 4) psqi_c7_std = 2.
IF (#s7 >= 5) psqi_c7_std = 3.
* Global score: only when all 7 components are available.
COMPUTE psqi_global_std = $SYSMIS.
IF (NMISS(psqi_c1_std, psqi_c2_std, psqi_c3_std, psqi_c4_std, psqi_c5_std, psqi_c6_std, psqi_c7_std) = 0)
   psqi_global_std = SUM(psqi_c1_std, psqi_c2_std, psqi_c3_std, psqi_c4_std, psqi_c5_std, psqi_c6_std, psqi_c7_std).
COMPUTE poor_sleep = $SYSMIS.
IF (NOT MISSING(psqi_global_std)) poor_sleep = (psqi_global_std > 5).
COMPUTE sleep_dur_cat = $SYSMIS.
IF (lsb_sle_psqi4 < 7) sleep_dur_cat = 1.
IF (lsb_sle_psqi4 >= 7 AND lsb_sle_psqi4 <= 9) sleep_dur_cat = 2.
IF (lsb_sle_psqi4 > 9) sleep_dur_cat = 3.
EXECUTE.

* =====================================================================.
* SECTION 10. PHYSICAL ACTIVITY - GPAQ (WHO analysis guide).
* Minutes/week per domain = days x minutes/day (source *_hours_to_minutes columns).
* Domain = 0 if answer "No" (code 2); missing if days not 1-7 or minutes/day > 960 (WHO cleaning).
* MET-min/week = 8 x vigorous + 4 x moderate + 4 x travel.
* =====================================================================.
DEFINE !pa_dom (out = !TOKENS(1) / yes = !TOKENS(1) / days = !TOKENS(1) / mins = !TOKENS(1))
COMPUTE !out = $SYSMIS.
IF (!yes = 2) !out = 0.
IF (!yes = 1 AND !days >= 1 AND !days <= 7 AND NOT MISSING(!mins) AND !mins <= 960) !out = !days * !mins.
!ENDDEFINE.
!pa_dom out=pa_vig_work_min yes=lsb_pa_gpaq1  days=lsb_pa_gpaq2  mins=gpaq3_hours_to_minutes.
!pa_dom out=pa_mod_work_min yes=lsb_pa_gpaq4  days=lsb_pa_gpaq5  mins=gpaq6_hours_to_minutes.
!pa_dom out=pa_travel_min   yes=lsb_pa_gpaq7  days=lsb_pa_gpaq8  mins=gpaq9a_hours_to_minutes.
!pa_dom out=pa_vig_rec_min  yes=lsb_pa_gpaq10 days=lsb_pa_gpaq11 mins=gpaq12_hours_to_minutes.
!pa_dom out=pa_mod_rec_min  yes=lsb_pa_gpaq13 days=lsb_pa_gpaq14 mins=gpaq15_hours_to_minutes.
COMPUTE pa_met_total = 8 * pa_vig_work_min + 4 * pa_mod_work_min + 4 * pa_travel_min + 8 * pa_vig_rec_min + 4 * pa_mod_rec_min.
COMPUTE pa_meets_who = $SYSMIS.
IF (NOT MISSING(pa_met_total)) pa_meets_who = (pa_met_total >= 600).
EXECUTE.

* =====================================================================.
* SECTION 11. COMPLETENESS / ANALYSIS-SET FLAGS.
* =====================================================================.
COMPUTE has_anthro = (NVALID(average_height, average_weight) = 2).
COMPUTE has_bp = (NVALID(pm_sbp, pm_dbp) = 2).
COMPUTE has_hb = NOT MISSING(hb_clean).
COMPUTE has_hba1c = NOT MISSING(hba1c_clean).
COMPUTE has_phq9 = NOT MISSING(psy_dep_phq9_score).
COMPUTE has_gse = NOT MISSING(psy_efficacy_gse_score).
COMPUTE has_psqi = NOT MISSING(psqi_global_std).
COMPUTE has_repro = NOT MISSING(rep_menstruation_age).
COMPUTE complete_core = (SUM(has_anthro, has_bp, has_hb, has_phq9, has_gse, has_psqi, has_repro) = 7).
* TRIAL ARM: add the randomisation group here when the allocation file is available, e.g.
*   MATCH FILES /FILE=* /TABLE='randomisation.sav' /BY record_id.
EXECUTE.

* =====================================================================.
* SECTION 12. FORMATS, VARIABLE AND VALUE LABELS.
* =====================================================================.
FORMATS hb_clean hba1c_clean (F8.1) edu_clean menarche_clean (F8.0) income_clean (F10.0)
  food_score_clean sedentary_min_day (F8.0) income_log (F8.2) sleep_eff_pct (F8.1)
  pa_vig_work_min pa_mod_work_min pa_travel_min pa_vig_rec_min pa_mod_rec_min pa_met_total (F8.0)
  flag_age flag_edu flag_menarche flag_bmi_low flag_height_low flag_hb flag_hba1c flag_sleep_eff
  flag_pulse flag_income flag_parity flag_name_mismatch flag_any
  age_grp edu_cat sli_cat fies_raw fies_cat food_insecure_any menarche_early living_cat abortion_any
  stillbirth_any ever_pregnant bmi_who bmi_asian underweight short_stature htn bp_cat anaemia
  anaemia_sev glyc_cat rbs_high diabetes_any phq9_cat depression phq9_item9_pos psqi_c1_std
  psqi_c2_std psqi_c3_std psqi_c4_std psqi_c5_std psqi_c6_std psqi_c7_std psqi_global_std
  poor_sleep sleep_dur_cat pa_meets_who has_anthro has_bp has_hb has_hba1c has_phq9 has_gse
  has_psqi has_repro complete_core (F4.0).

VARIABLE LABELS
  hb_clean 'Haemoglobin g/dL - cleaned (001054 corrected to 11.4; <2 or >20 missing)'
  hba1c_clean 'HbA1c % - cleaned (011066 corrected to 5.0; <3 or >20 missing)'
  edu_clean 'Education years - cleaned (>25 missing)'
  menarche_clean 'Age at menarche - cleaned (<8 or >20 missing)'
  income_clean 'Monthly household income - cleaned (<500 missing)'
  food_score_clean 'Food insecurity frequency score 0-24 - cleaned (all items blank = missing)'
  sedentary_min_day 'Sedentary time min/day (GPAQ Q16; >960 missing)'
  flag_age 'Check: age <18 or >35'
  flag_edu 'Check: education years >20'
  flag_menarche 'Check: menarche <9 or >18'
  flag_bmi_low 'Check: BMI <15'
  flag_height_low 'Check: height <140 cm'
  flag_hb 'Check: Hb <5 or >18 (raw)'
  flag_hba1c 'Check: HbA1c <3.5 or >15 (raw)'
  flag_sleep_eff 'Check: PSQI sleep efficiency >100% (hours slept > time in bed)'
  flag_pulse 'Check: pulse <45 or >130'
  flag_income 'Check: income <500'
  flag_parity 'Check: living children >=5 or stillbirths >=4'
  flag_name_mismatch 'Check: name/husband name differs from master in a data file'
  flag_any 'Any data-quality flag'
  age_grp 'Age group (years)'
  edu_cat 'Education level'
  income_log 'Log of monthly household income'
  sli_cat 'Standard of Living Index (NFHS-2 cut-offs)'
  fies_raw 'Food insecurity raw score: number of YES items (0-8)'
  fies_cat 'Food insecurity category (provisional FIES raw-score cut-offs)'
  food_insecure_any 'Any food insecurity (any item YES)'
  menarche_early 'Early menarche (<12 years)'
  living_cat 'Number of living children (grouped)'
  abortion_any 'Any previous abortion'
  stillbirth_any 'Any previous stillbirth'
  ever_pregnant 'Ever pregnant (any living child, abortion or stillbirth)'
  bmi_who 'BMI category - WHO'
  bmi_asian 'BMI category - WHO Asian cut-offs (23 / 27.5)'
  underweight 'Underweight (BMI <18.5)'
  short_stature 'Short stature (height <145 cm)'
  htn 'Hypertension (SBP >=140 or DBP >=90)'
  bp_cat 'Blood pressure category (ACC/AHA 2017)'
  anaemia 'Anaemia (Hb <12 g/dL, WHO non-pregnant women)'
  anaemia_sev 'Anaemia severity (WHO)'
  glyc_cat 'Glycaemic status by HbA1c (ADA)'
  rbs_high 'Random blood sugar >=200 mg/dL'
  diabetes_any 'Diabetes (HbA1c >=6.5 or RBS >=200)'
  phq9_cat 'PHQ-9 depression severity'
  depression 'Depression (PHQ-9 >=10)'
  phq9_item9_pos 'PHQ-9 item 9 positive (thoughts of self-harm) - SAFETY'
  psqi_c1_std 'PSQI C1 subjective sleep quality (standard scoring)'
  psqi_c2_std 'PSQI C2 sleep latency (standard scoring)'
  psqi_c3_std 'PSQI C3 sleep duration (standard scoring)'
  psqi_c4_std 'PSQI C4 habitual sleep efficiency (standard scoring)'
  psqi_c5_std 'PSQI C5 sleep disturbances (standard scoring)'
  psqi_c6_std 'PSQI C6 sleep medication (standard scoring)'
  psqi_c7_std 'PSQI C7 daytime dysfunction (standard scoring)'
  psqi_global_std 'PSQI global score 0-21 (standard scoring)'
  poor_sleep 'Poor sleep quality (PSQI global >5)'
  sleep_eff_pct 'Habitual sleep efficiency % (capped at 100)'
  sleep_dur_cat 'Sleep duration (PSQI Q4 hours)'
  pa_vig_work_min 'GPAQ vigorous work min/week'
  pa_mod_work_min 'GPAQ moderate work min/week'
  pa_travel_min 'GPAQ walking/cycling travel min/week'
  pa_vig_rec_min 'GPAQ vigorous recreation min/week'
  pa_mod_rec_min 'GPAQ moderate recreation min/week'
  pa_met_total 'GPAQ total physical activity MET-min/week'
  pa_meets_who 'Meets WHO physical activity recommendation (>=600 MET-min/week)'
  has_anthro 'Has height and weight'
  has_bp 'Has SBP and DBP'
  has_hb 'Has Hb (cleaned)'
  has_hba1c 'Has HbA1c (cleaned)'
  has_phq9 'Has PHQ-9 score'
  has_gse 'Has GSE score'
  has_psqi 'Has PSQI global score (standard)'
  has_repro 'Has reproductive history'
  complete_core 'Complete core data (anthro, BP, Hb, PHQ-9, GSE, PSQI, reproductive)'
  lsb_sle_psqi9 'PSQI Q6 subjective sleep quality (REDCap item 9)'
  lsb_sle_psqi6 'PSQI Q7 sleep medication (REDCap item 6)'
  lsb_sle_psqi7 'PSQI Q8 trouble staying awake (REDCap item 7)'
  lsb_sle_psqi8 'PSQI Q9 enthusiasm (REDCap item 8)'.

VALUE LABELS
  flag_age flag_edu flag_menarche flag_bmi_low flag_height_low flag_hb flag_hba1c flag_sleep_eff
  flag_pulse flag_income flag_parity flag_name_mismatch flag_any 0 'OK' 1 'Check'
 /food_insecure_any menarche_early abortion_any stillbirth_any ever_pregnant underweight short_stature
  htn anaemia rbs_high diabetes_any depression phq9_item9_pos poor_sleep pa_meets_who
  has_anthro has_bp has_hb has_hba1c has_phq9 has_gse has_psqi has_repro complete_core 0 'No' 1 'Yes'
 /age_grp 1 '<20' 2 '20-24' 3 '25-29' 4 '30+'
 /edu_cat 0 'None' 1 'Primary (1-5 y)' 2 'Secondary (6-10 y)' 3 'Higher secondary (11-12 y)' 4 'Graduate or above (13+ y)'
 /sli_cat 1 'Low (0-14)' 2 'Medium (15-24)' 3 'High (25-67)'
 /fies_cat 0 'Food secure (0)' 1 'Mild (1-3)' 2 'Moderate (4-6)' 3 'Severe (7-8)'
 /living_cat 0 'None' 1 'One' 2 'Two or more'
 /bmi_who 1 'Underweight <18.5' 2 'Normal 18.5-24.9' 3 'Overweight 25-29.9' 4 'Obese 30+'
 /bmi_asian 1 'Underweight <18.5' 2 'Normal 18.5-22.9' 3 'Overweight 23-27.4' 4 'Obese 27.5+'
 /bp_cat 1 'Normal' 2 'Elevated' 3 'Stage 1 hypertension' 4 'Stage 2 hypertension'
 /anaemia_sev 0 'None (>=12)' 1 'Mild (11-11.9)' 2 'Moderate (8-10.9)' 3 'Severe (<8)'
 /glyc_cat 1 'Normal (<5.7)' 2 'Prediabetes (5.7-6.4)' 3 'Diabetes (>=6.5)'
 /phq9_cat 0 'None (0-4)' 1 'Mild (5-9)' 2 'Moderate (10-14)' 3 'Moderately severe (15-19)' 4 'Severe (20-27)'
 /sleep_dur_cat 1 '<7 hours' 2 '7-9 hours' 3 '>9 hours'.

VARIABLE LEVEL age_grp edu_cat sli_cat fies_cat living_cat bmi_who bmi_asian bp_cat anaemia_sev
  glyc_cat phq9_cat sleep_dur_cat (ORDINAL).

* =====================================================================.
* SECTION 13. VERIFICATION - filled values per derived variable vs the reference counts
* (computed independently for the 20-Oct-2026 merged file). All rows should say PASS.
* If you later change a cleaning rule or the source data, FAIL is expected for affected rows.
* =====================================================================.
DATASET ACTIVATE analysis.
COMPUTE zz_one = 1.
DATASET DECLARE verify.
AGGREGATE /OUTFILE=verify /BREAK=zz_one
@AGG@.
DATASET ACTIVATE verify.
STRING zz_name (A8).
COMPUTE zz_name = 'found'.
EXECUTE.
DELETE VARIABLES zz_one.
FLIP /NEWNAMES=zz_name.
DELETE VARIABLES CASE_LBL.
COMPUTE check_no = $CASENUM.
FORMATS check_no (F4.0).
NUMERIC expected (F8.0).
STRING check_variable (A40).
EXECUTE.
@EXP@
STRING result (A4).
COMPUTE result = 'FAIL'.
IF (found = expected) result = 'PASS'.
EXECUTE.
SORT CASES BY result (A) check_no (A).
LIST check_no check_variable expected found result.
FREQUENCIES VARIABLES=result.
DATASET ACTIVATE analysis.
DELETE VARIABLES zz_one.
DATASET CLOSE verify.

* =====================================================================.
* SECTION 14. Summary tables of the derived variables.
* =====================================================================.
FREQUENCIES VARIABLES=age_grp edu_cat sli_cat fies_cat bmi_who bmi_asian bp_cat htn anaemia_sev
  glyc_cat diabetes_any phq9_cat depression phq9_item9_pos poor_sleep sleep_dur_cat pa_meets_who
  complete_core flag_any.
DESCRIPTIVES VARIABLES=hb_clean hba1c_clean edu_clean income_clean fies_raw psqi_global_std
  sleep_eff_pct pa_met_total sedentary_min_day /STATISTICS=MEAN STDDEV MIN MAX.

* =====================================================================.
* SECTION 15. Save.
* =====================================================================.
SAVE OUTFILE='HELTI_analysis_baseline.sav' /COMPRESSED.
SAVE TRANSLATE OUTFILE='HELTI_analysis_baseline.xlsx'
  /TYPE=XLS /VERSION=12 /MAP /FIELDNAMES VALUE=NAMES /REPLACE.
"""


def make_analysis_syntax(derived, name_ids, cd_folder="C:\\HELTI\\HELTI_merge_v5"):
    ids = [str(int(i)) for i in name_ids] or ["-1"]
    cols = list(derived.columns)
    s = TEMPLATE.replace("CD 'C:\\HELTI\\HELTI_merge_v5'.", f"CD '{cd_folder}'.")
    s = s.replace("@IDS@", "\n  ".join(" ".join(ids[i:i + 12]) for i in range(0, len(ids), 12)))
    s = s.replace("@AGG@", "\n".join(f"  /c{j:03d} = NU({c})" for j, c in enumerate(cols, 1)))
    s = s.replace("@EXP@", "\n".join(f"IF (check_no = {j}) expected = {int(derived[c].notna().sum())}.\nIF (check_no = {j}) check_variable = '{c}'." for j, c in enumerate(cols, 1)))
    return s
