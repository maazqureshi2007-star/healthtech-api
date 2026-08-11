from pathlib import Path
import pandas as pd


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

DATA_FILE = BASE_DIR / "data" / "knowledge_base.xlsx"


# ============================================================
# LOAD DATA
# ============================================================

if not DATA_FILE.exists():
    raise FileNotFoundError(
        f"Dataset not found: {DATA_FILE}"
    )

df = pd.read_excel(
    DATA_FILE,
    sheet_name=0,
    engine="openpyxl"
)

df = df.fillna("")


# ============================================================
# BASIC INFORMATION
# ============================================================

print("\n")
print("=" * 70)
print("HEALTHTECH DATASET ANALYSIS")
print("=" * 70)

print("\nDataset:")
print("File:", DATA_FILE)
print("Rows:", len(df))
print("Columns:", len(df.columns))


# ============================================================
# COLUMN ANALYSIS
# ============================================================

print("\n")
print("=" * 70)
print("COLUMN ANALYSIS")
print("=" * 70)

for column in df.columns:

    total = len(df)

    non_empty = (
        df[column]
        .astype(str)
        .str.strip()
        .ne("")
        .sum()
    )

    empty = total - non_empty

    percentage = (
        non_empty / total * 100
        if total > 0
        else 0
    )

    print(
        f"{column:<45} "
        f"filled={non_empty:<4} "
        f"empty={empty:<4} "
        f"coverage={percentage:.1f}%"
    )


# ============================================================
# DISEASE ID CHECK
# ============================================================

print("\n")
print("=" * 70)
print("DISEASE ID ANALYSIS")
print("=" * 70)

print(
    "Unique disease IDs:",
    df["disease_id"].nunique()
)

duplicate_ids = df[
    df["disease_id"].duplicated(keep=False)
]

print(
    "Duplicate disease IDs:",
    duplicate_ids["disease_id"].nunique()
)

if not duplicate_ids.empty:

    print("\nDuplicate IDs:")

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
# DISEASE NAME ANALYSIS
# ============================================================

print("\n")
print("=" * 70)
print("AYURVEDIC DISEASE NAME ANALYSIS")
print("=" * 70)

print(
    "Unique Ayurvedic names:",
    df["ayurvedic_name"].nunique()
)

print("\nMost common Ayurvedic names:")

print(
    df["ayurvedic_name"]
    .value_counts()
    .head(30)
    .to_string()
)


# ============================================================
# ENGLISH DISEASE ANALYSIS
# ============================================================

print("\n")
print("=" * 70)
print("ENGLISH DISEASE NAME ANALYSIS")
print("=" * 70)

print(
    "Unique English disease names:",
    df["disease_english"].nunique()
)

duplicates = (
    df["disease_english"]
    .value_counts()
)

print("\nRepeated English disease names:")

print(
    duplicates[duplicates > 1]
    .head(30)
    .to_string()
)


# ============================================================
# MEDICINE ANALYSIS
# ============================================================

print("\n")
print("=" * 70)
print("MEDICINE ANALYSIS")
print("=" * 70)

medicine_column = "ayurvedic_medicine_english"

medicine_values = (
    df[medicine_column]
    .astype(str)
    .str.strip()
)

print(
    "Rows containing medicine:",
    medicine_values.ne("").sum()
)

print(
    "Unique medicine strings:",
    medicine_values[
        medicine_values != ""
    ].nunique()
)

print("\nMost common medicine strings:")

print(
    medicine_values[
        medicine_values != ""
    ]
    .value_counts()
    .head(30)
    .to_string()
)


# ============================================================
# HERB ANALYSIS
# ============================================================

print("\n")
print("=" * 70)
print("AYURVEDIC HERB ANALYSIS")
print("=" * 70)

herb_column = "Ayurvedic Herbs"

herb_values = (
    df[herb_column]
    .astype(str)
    .str.strip()
)

print(
    "Rows containing herbs:",
    herb_values.ne("").sum()
)

print(
    "Unique herb strings:",
    herb_values[
        herb_values != ""
    ].nunique()
)

print("\nMost common herb strings:")

print(
    herb_values[
        herb_values != ""
    ]
    .value_counts()
    .head(30)
    .to_string()
)


# ============================================================
# FORMULATION ANALYSIS
# ============================================================

print("\n")
print("=" * 70)
print("FORMULATION ANALYSIS")
print("=" * 70)

formulation_values = (
    df["Formulation"]
    .astype(str)
    .str.strip()
)

print(
    "Rows containing formulations:",
    formulation_values.ne("").sum()
)

print(
    "Unique formulation strings:",
    formulation_values[
        formulation_values != ""
    ].nunique()
)

print("\nMost common formulation strings:")

print(
    formulation_values[
        formulation_values != ""
    ]
    .value_counts()
    .head(30)
    .to_string()
)


# ============================================================
# DOSHA ANALYSIS
# ============================================================

print("\n")
print("=" * 70)
print("DOSHA ANALYSIS")
print("=" * 70)

dosha_values = (
    df["doshas"]
    .astype(str)
    .str.strip()
)

print(
    "Rows containing dosha:",
    dosha_values.ne("").sum()
)

print("\nDosha values:")

print(
    dosha_values[
        dosha_values != ""
    ]
    .value_counts()
    .head(30)
    .to_string()
)


# ============================================================
# REFERENCE ANALYSIS
# ============================================================

print("\n")
print("=" * 70)
print("REFERENCE ANALYSIS")
print("=" * 70)

reference_values = (
    df["reference_english"]
    .astype(str)
    .str.strip()
)

print(
    "Rows containing references:",
    reference_values.ne("").sum()
)

print(
    "Unique reference strings:",
    reference_values[
        reference_values != ""
    ].nunique()
)


# ============================================================
# SAMPLE RECORDS
# ============================================================

print("\n")
print("=" * 70)
print("SAMPLE RECORDS")
print("=" * 70)

sample_columns = [
    "disease_id",
    "ayurvedic_name",
    "disease_english",
    "symptoms_english",
    "ayurvedic_medicine_english",
    "Ayurvedic Herbs",
    "Formulation",
    "doshas",
    "reference_english"
]

print(
    df[
        sample_columns
    ]
    .head(10)
    .to_string(index=False)
)


# ============================================================
# COMPLETE
# ============================================================

print("\n")
print("=" * 70)
print("ANALYSIS COMPLETE")
print("=" * 70)