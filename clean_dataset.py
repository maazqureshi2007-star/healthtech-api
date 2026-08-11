from pathlib import Path
import pandas as pd


# ============================================================
# PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

INPUT_FILE = BASE_DIR / "data" / "ayurvedic_data_raw.xlsx"
OUTPUT_FILE = BASE_DIR / "data" / "knowledge_base.xlsx"


# ============================================================
# KNOWLEDGE-BASE COLUMNS
# ============================================================

KNOWLEDGE_COLUMNS = [
    "disease_id",

    # Disease names
    "ayurvedic_name",
    "disease_english",
    "disease_hindi",
    "disease_gujarati",

    # Symptoms
    "symptoms_english",
    "symptoms_hindi",
    "symptoms_gujarati",

    # Treatment
    "treatment_english",
    "treatment_hindi",
    "treatment_gujarati",

    # Ayurvedic medicines
    "ayurvedic_medicine_english",
    "ayurvedic_medicine_hindi",
    "ayurvedic_medicine_gujarati",

    # References
    "reference_english",
    "reference_hindi",
    "reference_gujarati",

    # Clinical / knowledge information
    "diagnosis_and_tests",
    "symptom_severity",
    "duration_of_treatment",
    "risk_factors",
    "environmental_factors",
    "dietary_habits",
    "seasonal_variation",
    "age_group",
    "gender",

    # Ayurvedic concepts
    "doshas",
    "constitution_prakriti",

    # Recommendations
    "diet_and_lifestyle_recommendations",
    "yoga_and_physical_therapy",
    "prevention",
    "complications",

    # Herbs / formulations
    "Ayurvedic Herbs",
    "Formulation",

    # Medical information
    "Medical Intervention",
    "Prognosis"
]


# ============================================================
# LOAD RAW DATA
# ============================================================

if not INPUT_FILE.exists():
    raise FileNotFoundError(
        f"Input file not found: {INPUT_FILE}"
    )

print("Loading:", INPUT_FILE)

df = pd.read_excel(
    INPUT_FILE,
    sheet_name=0,
    engine="openpyxl"
)


# ============================================================
# CHECK REQUIRED COLUMNS
# ============================================================

missing_columns = [
    column
    for column in KNOWLEDGE_COLUMNS
    if column not in df.columns
]

if missing_columns:
    print("\nMissing columns:")
    for column in missing_columns:
        print("-", column)

    raise ValueError(
        "Required columns are missing from the Excel file."
    )


# ============================================================
# SELECT KNOWLEDGE COLUMNS
# ============================================================

clean_df = df[KNOWLEDGE_COLUMNS].copy()


# ============================================================
# CLEAN TEXT
# ============================================================

def clean_text(value):

    if pd.isna(value):
        return ""

    value = str(value)

    # Remove unnecessary spaces
    value = " ".join(value.split())

    return value.strip()


for column in clean_df.columns:

    if column != "disease_id":
        clean_df[column] = clean_df[column].apply(clean_text)


# ============================================================
# REMOVE EXACT DUPLICATE RECORDS
# ============================================================

before = len(clean_df)

clean_df = clean_df.drop_duplicates()

after = len(clean_df)

print("\nDuplicate rows removed:", before - after)


# ============================================================
# CHECK DUPLICATE DISEASE IDs
# ============================================================

duplicate_ids = clean_df[
    clean_df["disease_id"].duplicated(keep=False)
]

if not duplicate_ids.empty:

    print("\nWARNING: Duplicate disease IDs found:")

    print(
        duplicate_ids[
            [
                "disease_id",
                "ayurvedic_name",
                "disease_english"
            ]
        ].to_string(index=False)
    )


# ============================================================
# SORT BY DISEASE ID
# ============================================================

clean_df = clean_df.sort_values(
    by="disease_id"
)


# ============================================================
# SAVE CLEAN DATASET
# ============================================================

clean_df.to_excel(
    OUTPUT_FILE,
    index=False,
    engine="openpyxl"
)


# ============================================================
# SUMMARY
# ============================================================

print("\n========================================")
print("DATASET CLEANING COMPLETE")
print("========================================")

print("Original rows :", len(df))
print("Clean rows    :", len(clean_df))
print("Columns       :", len(clean_df.columns))
print("Output file   :", OUTPUT_FILE)

print("\nColumns:")
for column in clean_df.columns:
    print("-", column)