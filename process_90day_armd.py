import os
import sys
import pandas as pd

work_dir = r"C:\Users\samhi\OneDrive\Documents\sam\BITS\3-1\sop\sop code"
os.chdir(work_dir)

KEYS = ["anon_id", "pat_enc_csn_id_coded", "order_proc_id_coded"]

# Patient-level counts from the previous run (process_90day_armd.py)
OLD_REPORTED = {"med": 168331, "org": 24443, "both": 22462}


def log(msg=""):
    print(msg)
    sys.stdout.flush()


log("=" * 80)
log("PROCESSING ARMD DATASETS - CULTURE ORDER LEVEL (order_proc_id_coded)")
log("=" * 80)

# ----------------------------------------------------------------------------
# 1. Load
# ----------------------------------------------------------------------------
log("\n1. Loading files...")
df_med = pd.read_csv(
    "microbiology_cultures_prior_med.csv",
    usecols=KEYS + ["medication_name", "medication_time_to_culturetime"],
    dtype={"anon_id": "str", "pat_enc_csn_id_coded": "Int64",
           "order_proc_id_coded": "Int64", "medication_name": "str",
           "medication_time_to_culturetime": "float32"},
)
log(f"   Medications file: {len(df_med):,} rows loaded")

df_org = pd.read_csv(
    "microbiology_culture_prior_infecting_organism.csv",
    usecols=KEYS + ["prior_organism", "prior_infecting_organism_days_to_culutre"],
    dtype={"anon_id": "str", "pat_enc_csn_id_coded": "Int64",
           "order_proc_id_coded": "Int64", "prior_organism": "str",
           "prior_infecting_organism_days_to_culutre": "float32"},
)
log(f"   Organisms file: {len(df_org):,} rows loaded")

df_cohort = pd.read_csv(
    "microbiology_cultures_cohort.csv",
    usecols=["anon_id", "order_proc_id_coded"],
    dtype={"anon_id": "str", "order_proc_id_coded": "Int64"},
)
log(f"   Cohort file: {len(df_cohort):,} rows loaded")


def summarize(df, time_col, name_col, detail_col, total_col, distinct_col,
              details_col, label):
    """Filter to <=90 days and aggregate per culture order."""
    # Warn on nulls in the time column
    n_null_time = int(df[time_col].isna().sum())
    if n_null_time:
        log(f"   WARNING: Skipping {n_null_time:,} rows with null {time_col}")
    df = df.dropna(subset=[time_col])

    # Warn on nulls in the grouping keys / name column
    key_null_mask = df[KEYS].isna().any(axis=1) | df[name_col].isna()
    n_null_key = int(key_null_mask.sum())
    if n_null_key:
        log(f"   WARNING: Skipping {n_null_key:,} rows with null {KEYS} or {name_col}")
    df = df[~key_null_mask]

    df = df[df[time_col] <= 90].copy()
    log(f"   Filtered to {len(df):,} rows ({time_col} <= 90)")

    df[time_col] = df[time_col].astype("int32")
    df[detail_col] = df[name_col] + " (" + df[time_col].astype(str) + "d)"

    out = (
        df.groupby(KEYS, sort=False)
        .agg(**{
            total_col: (name_col, "size"),
            distinct_col: (name_col, "nunique"),
            details_col: (detail_col, ", ".join),
        })
        .reset_index()
    )
    out = out.sort_values(["anon_id", "pat_enc_csn_id_coded", "order_proc_id_coded"],
                          kind="stable").reset_index(drop=True)
    log(f"   Created summary for {len(out):,} culture orders "
        f"({out['anon_id'].nunique():,} patients) [{label}]")
    return out, set(df["anon_id"].unique())


# ----------------------------------------------------------------------------
# 2. Medications
# ----------------------------------------------------------------------------
log("\n2. Processing medications file...")
df_med_summary, old_style_med_patients = summarize(
    df_med, "medication_time_to_culturetime", "medication_name", "drug_detail_str",
    "total_exposures_within_90d", "distinct_drugs_within_90d", "drug_details",
    "medication",
)
del df_med

# ----------------------------------------------------------------------------
# 3. Organisms
# ----------------------------------------------------------------------------
log("\n3. Processing organisms file...")
df_org_summary, old_style_org_patients = summarize(
    df_org, "prior_infecting_organism_days_to_culutre", "prior_organism",
    "org_detail_str", "total_organism_records_within_90d",
    "distinct_organisms_within_90d", "organism_details", "organism",
)
del df_org

# ----------------------------------------------------------------------------
# 4. Validation against cohort
# ----------------------------------------------------------------------------
log("\n4. Validating against cohort file...")
cohort_patients = set(df_cohort["anon_id"].unique())
cohort_orders = set(df_cohort["order_proc_id_coded"].dropna().unique())
for label, df_sum in (("Medications", df_med_summary), ("Organisms", df_org_summary)):
    pats = set(df_sum["anon_id"].unique())
    ords = set(df_sum["order_proc_id_coded"].unique())
    log(f"   {label}: anon_id   match {len(pats & cohort_patients):,} / "
        f"no match {len(pats - cohort_patients):,}")
    log(f"   {label}: order_id  match {len(ords & cohort_orders):,} / "
        f"no match {len(ords - cohort_orders):,}")

# ----------------------------------------------------------------------------
# 5. Combined (same culture order has both medication AND organism records)
# ----------------------------------------------------------------------------
log("\n5. Building combined shortlist...")
df_combined = df_med_summary.merge(df_org_summary, on=KEYS, how="inner")
df_combined = df_combined.sort_values(
    ["anon_id", "pat_enc_csn_id_coded", "order_proc_id_coded"], kind="stable"
).reset_index(drop=True)
df_combined = df_combined[KEYS + [
    "total_exposures_within_90d", "distinct_drugs_within_90d", "drug_details",
    "total_organism_records_within_90d", "distinct_organisms_within_90d",
    "organism_details",
]]
log(f"   {len(df_combined):,} culture orders in both")

# ----------------------------------------------------------------------------
# 6. Save
# ----------------------------------------------------------------------------
log("\n6. Saving Excel files...")
outputs = {
    "prior_medication_90day_shortlist.xlsx": df_med_summary,
    "prior_organism_90day_shortlist.xlsx": df_org_summary,
    "both_conditions_90day_shortlist.xlsx": df_combined,
}
for fname, df in outputs.items():
    df.to_excel(os.path.join(work_dir, fname), index=False, engine="openpyxl")
    log(f"   Saved {fname} ({len(df):,} rows)")

# ----------------------------------------------------------------------------
# 7. Final summary: old (patient-level) vs new (culture-order-level)
# ----------------------------------------------------------------------------
# Old-run numbers recomputed with the old logic (unique anon_id with any record <= 90d)
old_recomputed = {
    "med": len(old_style_med_patients),
    "org": len(old_style_org_patients),
    "both": len(old_style_med_patients & old_style_org_patients),
}
new_orders = {"med": len(df_med_summary), "org": len(df_org_summary),
              "both": len(df_combined)}
new_patients = {"med": df_med_summary["anon_id"].nunique(),
                "org": df_org_summary["anon_id"].nunique(),
                "both": df_combined["anon_id"].nunique()}
names = {"med": "prior_medication_90day_shortlist.xlsx",
         "org": "prior_organism_90day_shortlist.xlsx",
         "both": "both_conditions_90day_shortlist.xlsx"}

log("\n" + "=" * 100)
log("FINAL SUMMARY: OLD (patient-level) vs NEW (culture-order-level)")
log("=" * 100)
log(f"{'File':<40}{'OLD patients':>14}{'(recomputed)':>14}{'NEW patients':>14}{'NEW orders':>14}")
log("-" * 100)
for k in ("med", "org", "both"):
    log(f"{names[k]:<40}{OLD_REPORTED[k]:>14,}{old_recomputed[k]:>14,}"
        f"{new_patients[k]:>14,}{new_orders[k]:>14,}")
log("-" * 100)
log("Note: for 'both', OLD = patients in both lists (any order); NEW = patients having at least")
log("one single culture order with BOTH medication and organism records within 90 days.")
log("=" * 100)
