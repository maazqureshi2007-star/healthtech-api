from pathlib import Path
import pandas as pd


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

DATA_FILE = (
    BASE_DIR
    / "data"
    / "normalized_knowledge_base.xlsx"
)


# ============================================================
# LOAD DATA
# ============================================================

if not DATA_FILE.exists():
    raise FileNotFoundError(
        f"Normalized dataset not found: {DATA_FILE}"
    )


df = pd.read_excel(
    DATA_FILE,
    sheet_name=0,
    engine="openpyxl"
)


df = df.fillna("")


# ============================================================
# HEADER
# ============================================================

print()
print("=" * 75)
print("HEALTHTECH NORMALIZED DATASET VALIDATION")
print("=" * 75)


# ============================================================
# BASIC CHECK
# ============================================================

print()
print("DATASET")
print("-" * 75)

print("Rows:", len(df))
print("Columns:", len(df.columns))


# ============================================================
# REQUIRED COLUMNS
# ============================================================

required_columns = [
    "disease_id",
    "ayurvedic_name",
    "disease_english",
    "symptoms_english",
    "treatment_english",
    "ayurvedic_medicine_english",
    "Ayurvedic Herbs",
    "Formulation",
    "doshas",
    "reference_english",

    "disease_english_search",
    "symptoms_search",
    "medicine_search",
    "herbs_search",
    "formulation_search",
    "doshas_search",

    "doshas_normalized",
    "herbs_normalized",
    "medicine_terms_normalized",

    "reference_source",
    "reference_section",

    "formulation_raw",
    "search_document",
]


print()
print("REQUIRED COLUMN CHECK")
print("-" * 75)

missing_columns = []

for column in required_columns:

    if column in df.columns:
        print(f"OK   {column}")
    else:
        print(f"FAIL {column}")
        missing_columns.append(column)


# ============================================================
# DISEASE ID CHECK
# ============================================================

print()
print("DISEASE ID CHECK")
print("-" * 75)

total_ids = len(df)

unique_ids = df["disease_id"].nunique()

duplicate_ids = (
    df["disease_id"]
    .duplicated()
    .sum()
)

empty_ids = (
    df["disease_id"]
    .astype(str)
    .str.strip()
    .eq("")
    .sum()
)

print("Total IDs:", total_ids)
print("Unique IDs:", unique_ids)
print("Duplicate IDs:", duplicate_ids)
print("Empty IDs:", empty_ids)


# ============================================================
# ENGLISH DISEASE CHECK
# ============================================================

print()
print("DISEASE NAME CHECK")
print("-" * 75)

empty_disease_names = (
    df["disease_english"]
    .astype(str)
    .str.strip()
    .eq("")
    .sum()
)

print(
    "Empty English disease names:",
    empty_disease_names
)

print(
    "Unique English disease names:",
    df["disease_english"].nunique()
)


# ============================================================
# SEARCH FIELD CHECK
# ============================================================

print()
print("SEARCH FIELD CHECK")
print("-" * 75)

search_columns = [
    "disease_english_search",
    "ayurvedic_name_search",
    "symptoms_search",
    "treatment_search",
    "medicine_search",
    "herbs_search",
    "formulation_search",
    "doshas_search",
    "reference_search",
    "search_document",
]

for column in search_columns:

    empty = (
        df[column]
        .astype(str)
        .str.strip()
        .eq("")
        .sum()
    )

    populated = len(df) - empty

    print(
        f"{column:<35}"
        f" populated={populated:<4}"
        f" empty={empty:<4}"
    )


# ============================================================
# DOSHA CHECK
# ============================================================

print()
print("DOSHA NORMALIZATION CHECK")
print("-" * 75)

for index, row in df.head(10).iterrows():

    print()
    print("Disease:", row["disease_english"])
    print("Original:", row["doshas"])
    print("Normalized:", row["doshas_normalized"])


# ============================================================
# HERB CHECK
# ============================================================

print()
print("HERB NORMALIZATION CHECK")
print("-" * 75)

for index, row in df.head(10).iterrows():

    print()
    print("Disease:", row["disease_english"])
    print("Original:", row["Ayurvedic Herbs"])
    print("Normalized:", row["herbs_normalized"])


# ============================================================
# MEDICINE CHECK
# ============================================================

print()
print("MEDICINE NORMALIZATION CHECK")
print("-" * 75)

for index, row in df.head(10).iterrows():

    print()
    print("Disease:", row["disease_english"])
    print(
        "Original:",
        row["ayurvedic_medicine_english"]
    )
    print(
        "Normalized:",
        row["medicine_terms_normalized"]
    )


# ============================================================
# REFERENCE CHECK
# ============================================================

print()
print("REFERENCE CHECK")
print("-" * 75)

for index, row in df.head(10).iterrows():

    print()
    print("Disease:", row["disease_english"])
    print(
        "Original:",
        row["reference_english"]
    )
    print(
        "Source:",
        row["reference_source"]
    )
    print(
        "Section:",
        row["reference_section"]
    )


# ============================================================
# SEARCH DOCUMENT CHECK
# ============================================================

print()
print("SEARCH DOCUMENT CHECK")
print("-" * 75)

for index, row in df.head(5).iterrows():

    print()
    print("Disease:", row["disease_english"])

    print(
        "Search document:"
    )

    print(
        row["search_document"]
    )


# ============================================================
# FINAL VALIDATION
# ============================================================

print()
print("=" * 75)
print("VALIDATION SUMMARY")
print("=" * 75)

errors = []

if missing_columns:
    errors.append(
        f"Missing columns: {missing_columns}"
    )

if duplicate_ids > 0:
    errors.append(
        f"Duplicate disease IDs: {duplicate_ids}"
    )

if empty_ids > 0:
    errors.append(
        f"Empty disease IDs: {empty_ids}"
    )

if empty_disease_names > 0:
    errors.append(
        f"Empty disease names: {empty_disease_names}"
    )


if errors:

    print()
    print("VALIDATION FAILED")

    for error in errors:
        print("ERROR:", error)

else:

    print()
    print("VALIDATION PASSED")
    print()
    print(
        "The normalized dataset is structurally ready "
        "for the API layer."
    )

print()