from pathlib import Path
import pandas as pd
import re
import unicodedata

from main import clean_value


# ============================================================
# PATH CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

INPUT_FILE = BASE_DIR / "data" / "knowledge_base.xlsx"
OUTPUT_FILE = BASE_DIR / "data" / "normalized_knowledge_base.xlsx"


# ============================================================
# LOAD DATA
# ============================================================

if not INPUT_FILE.exists():
    raise FileNotFoundError(
        f"Knowledge base not found: {INPUT_FILE}"
    )

df = pd.read_excel(
    INPUT_FILE,
    sheet_name=0,
    engine="openpyxl"
)

# Clean Excel column names
def clean_column_name(column):
    column = str(column)
    column = unicodedata.normalize("NFKC", column)
    column = column.replace("\ufeff", "")
    column = column.replace("\xa0", " ")
    column = re.sub(r"[\u200B-\u200D\u2060]", "", column)
    column = re.sub(r"\s+", " ", column)
    return column.strip()

df.columns = [clean_column_name(column) for column in df.columns]


# ============================================================
# FIX DATASET COLUMN SPELLING
# ============================================================

COLUMN_RENAMES = {
    "disease_gujarati": "disease_gujarati",
    "symptoms_gujarati": "symptoms_gujarati",
    "treatment_gujarati": "treatment_gujarati",
    "ayurvedic_medicine_gujarati": "ayurvedic_medicine_gujarati",
    "reference_gujarati": "reference_gujarati",
}

df.rename(columns=COLUMN_RENAMES, inplace=True)


df = df.fillna("")


# ============================================================
# TEXT NORMALIZATION
# ============================================================

def normalize_text(value):
    """
    Normalize text for searching/comparison.

    This does NOT change the original meaning.
    """

    if pd.isna(value):
        return ""

    text = str(value).strip()

    # Normalize whitespace
    text = re.sub(r"\s+", " ", text)

    # Normalize common separators
    text = text.replace("，", ",")
    text = text.replace(";", ",")

    return text.strip()


def normalize_search_text(value):
    """
    Create a lowercase searchable representation.
    """

    text = normalize_text(value)

    text = text.lower()

    # Remove unnecessary punctuation
    text = re.sub(r"[^\w\s,.-]", " ", text)

    # Normalize whitespace again
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def split_comma_values(value):
    """
    Safely split comma-separated values.

    Used only for fields where comma-separated
    values represent multiple concepts.
    """

    text = normalize_text(value)

    if not text:
        return []

    parts = [
        item.strip()
        for item in text.split(",")
        if item.strip()
    ]

    return parts


def unique_preserve_order(items):
    """
    Remove duplicates without changing order.
    """

    result = []

    seen = set()

    for item in items:

        key = item.lower().strip()

        if key and key not in seen:

            result.append(item.strip())
            seen.add(key)

    return result


# ============================================================
# BASIC TEXT NORMALIZATION
# ============================================================

TEXT_COLUMNS = [
    "ayurvedic_name",
    "disease_english",
    "disease_hindi",
    "disease_gujarati",

    "symptoms_english",
    "symptoms_hindi",
    "symptoms_gujarati",

    "treatment_english",
    "treatment_hindi",
    "treatment_gujarati",

    "ayurvedic_medicine_english",
    "ayurvedic_medicine_hindi",
    "ayurvedic_medicine_gujarati",

    "reference_english",
    "reference_hindi",
    "reference_gujarati",

    "diagnosis_and_tests",
    "symptom_severity",
    "duration_of_treatment",
    "risk_factors",
    "environmental_factors",
    "dietary_habits",
    "seasonal_variation",
    "age_group",
    "gender",

    "doshas",
    "constitution_prakriti",

    "diet_and_lifestyle_recommendations",
    "yoga_and_physical_therapy",
    "prevention",
    "complications",

    "Ayurvedic Herbs",
    "Formulation",
    "Medical Intervention",
    "Prognosis"
]


for column in TEXT_COLUMNS:

    if column in df.columns:

        df[column] = df[column].apply(
            normalize_text
        )


# ============================================================
# SEARCHABLE FIELDS
# ============================================================

df["disease_english_search"] = (
    df["disease_english"]
    .apply(normalize_search_text)
)

df["disease_gujarati_search"] = (
    df["disease_gujarati"]
    .apply(normalize_search_text)
)

df["ayurvedic_name_search"] = (
    df["ayurvedic_name"]
    .apply(normalize_search_text)
)

df["symptoms_search"] = (
    df["symptoms_english"]
    .apply(normalize_search_text)
)

df["treatment_search"] = (
    df["treatment_english"]
    .apply(normalize_search_text)
)

df["medicine_search"] = (
    df["ayurvedic_medicine_english"]
    .apply(normalize_search_text)
)

df["herbs_search"] = (
    df["Ayurvedic Herbs"]
    .apply(normalize_search_text)
)

df["formulation_search"] = (
    df["Formulation"]
    .apply(normalize_search_text)
)

df["doshas_search"] = (
    df["doshas"]
    .apply(normalize_search_text)
)


# ============================================================
# NORMALIZED DOSHAS
# ============================================================

df["doshas_normalized"] = df["doshas"].apply(
    lambda value: unique_preserve_order(
        split_comma_values(value)
    )
)


# ============================================================
# NORMALIZED HERBS
# ============================================================

df["herbs_normalized"] = df["Ayurvedic Herbs"].apply(
    lambda value: unique_preserve_order(
        split_comma_values(value)
    )
)


# ============================================================
# NORMALIZED MEDICINE TERMS
# ============================================================

df["medicine_terms_normalized"] = (
    df["ayurvedic_medicine_english"]
    .apply(
        lambda value: unique_preserve_order(
            split_comma_values(value)
        )
    )
)


# ============================================================
# REFERENCE NORMALIZATION
# ============================================================

df["reference_search"] = (
    df["reference_english"]
    .apply(normalize_search_text)
)


# ============================================================
# REFERENCE SOURCE EXTRACTION
# ============================================================

def extract_reference_source(value):

    text = normalize_text(value)

    if not text:
        return ""

    lowered = text.lower()

    if "charaka samhita" in lowered:
        return "Charaka Samhita"

    if "sushruta samhita" in lowered:
        return "Sushruta Samhita"

    if "ashtanga hridaya" in lowered:
        return "Ashtanga Hridaya"

    if "ashtanga sangraha" in lowered:
        return "Ashtanga Sangraha"

    return text


df["reference_source"] = (
    df["reference_english"]
    .apply(extract_reference_source)
)


# ============================================================
# REFERENCE SECTION
# ============================================================

def extract_reference_section(value):

    text = normalize_text(value)

    if not text:
        return ""

    parts = [
        part.strip()
        for part in text.split(",")
        if part.strip()
    ]

    if len(parts) >= 2:

        return ", ".join(parts[1:])

    return ""


df["reference_section"] = (
    df["reference_english"]
    .apply(extract_reference_section)
)


# ============================================================
# FORMULATION RAW PRESERVATION
# ============================================================

# IMPORTANT:
# We deliberately preserve the original formulation text.
#
# We do NOT attempt to automatically interpret:
#
# - dosage
# - ingredients
# - preparation method
# - anupana
#
# because doing that incorrectly could create false medical data.

df["formulation_raw"] = df["Formulation"]


# ============================================================
# RECORD TYPE
# ============================================================

df["record_type"] = "ayurvedic_knowledge"


# ============================================================
# API SEARCH TEXT
# ============================================================

def build_search_document(row):

    fields = [
        row.get("ayurvedic_name", ""),
        row.get("disease_english", ""),
        row.get("symptoms_english", ""),
        row.get("treatment_english", ""),
        row.get("ayurvedic_medicine_english", ""),
        row.get("Ayurvedic Herbs", ""),
        row.get("Formulation", ""),
        row.get("doshas", ""),
        row.get("reference_english", "")
    ]

    fields = [
        normalize_text(value)
        for value in fields
        if normalize_text(value)
    ]

    return " | ".join(fields)


df["search_document"] = df.apply(
    build_search_document,
    axis=1
)


# ============================================================
# REMOVE EXACT DUPLICATES
# ============================================================

before = len(df)

df = df.drop_duplicates(
    subset=["disease_id"]
)

after = len(df)

print(
    "Duplicate disease records removed:",
    before - after
)


# ============================================================
# SORT
# ============================================================

df = df.sort_values(
    by="disease_id"
)


# ============================================================
# SAVE
# ============================================================

df.to_excel(
    OUTPUT_FILE,
    index=False,
    engine="openpyxl"
)


# ============================================================
# SUMMARY
# ============================================================

print()
print("=" * 70)
print("NORMALIZATION COMPLETE")
print("=" * 70)

print("Input file:")
print(INPUT_FILE)

print()

print("Output file:")
print(OUTPUT_FILE)

print()

print("Records:", len(df))
print("Columns:", len(df.columns))

print()

print("New normalized fields:")

NEW_COLUMNS = [
    "disease_english_search",
    "disease_gujarati_search",
    "ayurvedic_name_search",
    "symptoms_search",
    "treatment_search",
    "medicine_search",
    "herbs_search",
    "formulation_search",
    "doshas_search",
    "doshas_normalized",
    "herbs_normalized",
    "medicine_terms_normalized",
    "reference_search",
    "reference_source",
    "reference_section",
    "formulation_raw",
    "record_type",
    "search_document"
]

for column in NEW_COLUMNS:

    print("-", column)

print()
print("Normalization finished successfully.")