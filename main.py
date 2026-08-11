from pathlib import Path
from typing import Optional

import pandas as pd
from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel


# ============================================================
# APPLICATION
# ============================================================

app = FastAPI(
    title="HealthTech Ayurvedic Knowledge API",
    description=(
        "API for Ayurvedic disease, medicine, formulation, symptom, "
        "dosha and reference data."
    ),
    version="1.0.0",
)

# Frontend development origin used by the current project.
# Keep this permissive for local SIH integration; restrict it before production.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ============================================================
# DATA CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "data" / "normalized_knowledge_base.xlsx"

df: Optional[pd.DataFrame] = None


# ============================================================
# SEARCH ALIASES
# ============================================================

SEARCH_ALIASES = {
    # Gujarati
    "મધુપ્રમેહ": "મધુમેહ",
}


# ============================================================
# DATASET COLUMN HELPERS
# ============================================================

# Some older files used "gujrati". The current normalized dataset uses
# "gujarati". At startup we convert the old spelling only when necessary.
GUJARATI_LEGACY_ALIASES = {
    "disease_gujrati": "disease_gujarati",
    "symptoms_gujrati": "symptoms_gujarati",
    "treatment_gujrati": "treatment_gujarati",
    "medicine_gujrati": "medicine_gujarati",
    "ayurvedic_medicine_gujrati": "ayurvedic_medicine_gujarati",
    "herbs_gujrati": "herbs_gujarati",
    "formulation_gujrati": "formulation_gujarati",
    "doshas_gujrati": "doshas_gujarati",
    "reference_gujrati": "reference_gujarati",
}


def normalize_columns(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Normalize whitespace and safely migrate old Gujarati column spellings."""
    dataframe = dataframe.copy()

    dataframe.columns = (
        dataframe.columns.astype(str)
        .str.strip()
    )

    rename_map = {
        old: new
        for old, new in GUJARATI_LEGACY_ALIASES.items()
        if old in dataframe.columns and new not in dataframe.columns
    }

    if rename_map:
        dataframe = dataframe.rename(columns=rename_map)

    return dataframe


def load_data() -> pd.DataFrame:
    """Load and normalize the Excel knowledge base."""
    global df

    if not DATA_FILE.exists():
        raise FileNotFoundError(
            f"Excel file not found: {DATA_FILE}"
        )

    data = pd.read_excel(
        DATA_FILE,
        sheet_name=0,
        engine="openpyxl",
    )

    data = normalize_columns(data)

    # Replace NaN/NaT with empty strings so JSON responses are safe.
    data = data.fillna("")

    # Keep one row per disease ID when the ID exists.
    if "disease_id" in data.columns:
        data = data.drop_duplicates(
            subset=["disease_id"],
            keep="first",
        ).reset_index(drop=True)

    df = data

    print(
        f"Loaded dataset: {DATA_FILE.name} | "
        f"rows={len(df)} | columns={len(df.columns)}"
    )

    return df


# Load when API starts.
load_data()


# ============================================================
# GENERIC HELPERS
# ============================================================

def clean_value(value, default=""):
    """Convert pandas values into JSON-safe strings."""
    if value is None:
        return default

    try:
        if pd.isna(value):
            return default
    except (TypeError, ValueError):
        pass

    return str(value).strip()


def normalize_text(value) -> str:
    """Normalize text for case-insensitive multilingual substring matching."""
    return clean_value(value).casefold().strip()


def get_row_value(row: pd.Series, column: str, default=""):
    """Safely read a column that may not exist in every dataset version."""
    if column not in row.index:
        return default
    return clean_value(row[column], default)


def row_to_dict(row: pd.Series) -> dict:
    """Convert an entire dataframe row to a JSON-safe dictionary."""
    return {
        str(column): clean_value(row[column])
        for column in df.columns
    }


def first_available(row: pd.Series, columns: list[str], default=""):
    """Return the first non-empty value among candidate columns."""
    for column in columns:
        value = get_row_value(row, column, "")
        if value:
            return value
    return default


def disease_id_matches(series: pd.Series, disease_id: int) -> pd.Series:
    """Match numeric or string disease IDs safely."""
    target = str(disease_id).strip()
    return series.astype(str).str.strip() == target


def unique_results_by_disease_id(items: list[dict], limit: int) -> list[dict]:
    """Remove duplicate disease IDs while preserving ranking order."""
    results = []
    seen = set()

    for item in items:
        disease_id = clean_value(item.get("disease_id"))

        # If an ID is missing, do not collapse unrelated records into one.
        if disease_id and disease_id in seen:
            continue

        if disease_id:
            seen.add(disease_id)

        results.append(item)

        if len(results) >= limit:
            break

    return results


def build_search_groups() -> dict:
    """
    Search fields grouped by relevance.

    A group contributes its score only once even if several columns in
    the same group contain the query.
    """
    return {
        "disease": {
            "columns": [
                "disease_english",
                "disease_hindi",
                "disease_gujarati",
                "disease_english_search",
                "disease_hindi_search",
                "disease_gujarati_search",
            ],
            "score": 10,
        },
        "ayurvedic_name": {
            "columns": [
                "ayurvedic_name",
                "ayurvedic_name_search",
            ],
            "score": 8,
        },
        "symptoms": {
            "columns": [
                "symptoms_english",
                "symptoms_hindi",
                "symptoms_gujarati",
                "symptoms_search",
            ],
            "score": 6,
        },
        "treatment": {
            "columns": [
                "treatment_english",
                "treatment_hindi",
                "treatment_gujarati",
                "treatment_search",
            ],
            "score": 4,
        },
        "medicine": {
            "columns": [
                "medicine_english",
                "medicine_hindi",
                "medicine_gujarati",
                "ayurvedic_medicine_english",
                "ayurvedic_medicine_hindi",
                "ayurvedic_medicine_gujarati",
                "medicine_search",
            ],
            "score": 3,
        },
        "herbs": {
            "columns": [
                "herbs_english",
                "herbs_hindi",
                "herbs_gujarati",
                "Ayurvedic Herbs",
                "herbs_search",
            ],
            "score": 3,
        },
        "formulation": {
            "columns": [
                "formulation_english",
                "formulation_hindi",
                "formulation_gujarati",
                "Formulation",
                "formulation_search",
            ],
            "score": 2,
        },
        "dosha": {
            "columns": [
                "doshas",
                "doshas_english",
                "doshas_hindi",
                "doshas_gujarati",
                "doshas_normalized",
                "doshas_search",
            ],
            "score": 2,
        },
        "reference": {
            "columns": [
                "reference_english",
                "reference_hindi",
                "reference_gujarati",
                "reference_source",
                "reference_section",
                "reference_search",
            ],
            "score": 1,
        },
        "search_document": {
            "columns": ["search_document"],
            "score": 1,
        },
    }


def make_search_result(row: pd.Series, score: int, matched_fields: list[str]) -> dict:
    """Create the compact search-result object used by the frontend."""
    return {
        "disease_id": first_available(row, ["disease_id"]),
        "ayurvedic_name": first_available(row, ["ayurvedic_name"]),
        "disease_english": first_available(row, ["disease_english"]),
        "disease_hindi": first_available(row, ["disease_hindi"]),
        "disease_gujarati": first_available(row, ["disease_gujarati"]),
        "symptoms": first_available(
            row,
            ["symptoms_english", "symptoms_hindi", "symptoms_gujarati"],
        ),
        "medicines": first_available(
            row,
            [
                "ayurvedic_medicine_english",
                "medicine_english",
                "ayurvedic_medicine_hindi",
                "medicine_hindi",
                "ayurvedic_medicine_gujarati",
                "medicine_gujarati",
            ],
        ),
        "herbs": first_available(
            row,
            ["Ayurvedic Herbs", "herbs_english", "herbs_hindi", "herbs_gujarati"],
        ),
        "formulation": first_available(
            row,
            [
                "Formulation",
                "formulation_english",
                "formulation_hindi",
                "formulation_gujarati",
            ],
        ),
        "doshas": first_available(
            row,
            ["doshas", "doshas_normalized", "doshas_english", "doshas_hindi"],
        ),
        "reference": first_available(
            row,
            ["reference_english", "reference_hindi", "reference_gujarati"],
        ),
        "matched_fields": matched_fields,
        "score": score,
    }


# ============================================================
# ROOT
# ============================================================

@app.get("/")
def root():
    return {
        "name": "HealthTech Ayurvedic Knowledge API",
        "version": "1.0.0",
        "status": "running",
        "dataset_rows": len(df),
        "dataset_columns": len(df.columns),
    }


# ============================================================
# HEALTH CHECK
# ============================================================

@app.get("/api/v1/health")
def health():
    return {
        "status": "ok",
        "dataset_loaded": df is not None and not df.empty,
        "dataset_file": DATA_FILE.name,
        "records": len(df),
        "columns": len(df.columns),
    }


# ============================================================
# DATASET INFORMATION
# ============================================================

@app.get("/api/v1/dataset")
def dataset_info():
    return {
        "file": DATA_FILE.name,
        "rows": len(df),
        "columns": list(df.columns),
    }


# ============================================================
# GET ALL DISEASES
# ============================================================

@app.get("/api/v1/diseases")
def get_diseases(
    limit: int = Query(20, ge=1, le=500),
    offset: int = Query(0, ge=0),
):
    records = df.iloc[offset: offset + limit]

    data = []
    for _, row in records.iterrows():
        data.append({
            "disease_id": first_available(row, ["disease_id"]),
            "ayurvedic_name": first_available(row, ["ayurvedic_name"]),
            "disease_english": first_available(row, ["disease_english"]),
            "disease_hindi": first_available(row, ["disease_hindi"]),
            "disease_gujarati": first_available(row, ["disease_gujarati"]),
            "symptoms_english": first_available(row, ["symptoms_english"]),
            "ayurvedic_medicine_english": first_available(
                row,
                ["ayurvedic_medicine_english", "medicine_english"],
            ),
            "reference_english": first_available(
                row,
                ["reference_english"],
            ),
        })

    return {
        "status": "success",
        "count": len(data),
        "total": len(df),
        "offset": offset,
        "limit": limit,
        "data": data,
    }


# ============================================================
# SEARCH
# ============================================================

@app.get("/api/v1/search")
def search(
    q: str = Query(..., min_length=2),
    limit: int = Query(10, ge=1, le=50),
):
    """
    Search the Ayurvedic knowledge base across English, Hindi and Gujarati.

    Ranking:
    disease > Ayurvedic name > symptoms > treatment > medicine/herbs >
    formulation/dosha > reference.
    """
    if df is None or df.empty:
        raise HTTPException(
            status_code=500,
            detail="Knowledge base is not loaded",
        )

    original_query = clean_value(q).strip()
    query = normalize_text(original_query)

    if not query:
        raise HTTPException(
            status_code=400,
            detail="Search query cannot be empty",
        )

    query = SEARCH_ALIASES.get(query, query)

    search_groups = build_search_groups()
    scored = []

    for _, row in df.iterrows():
        score = 0
        matched_fields = []

        for group_name, group in search_groups.items():
            existing_columns = [
                column
                for column in group["columns"]
                if column in df.columns
            ]

            group_matched = False

            for column in existing_columns:
                value = normalize_text(row[column])

                if value and query in value:
                    group_matched = True
                    break

            if group_matched:
                score += group["score"]
                matched_fields.append(group_name)

        if score > 0:
            scored.append({
                "score": score,
                "matched_fields": matched_fields,
                "row": row,
            })

    scored.sort(
        key=lambda item: item["score"],
        reverse=True,
    )

    raw_results = [
        make_search_result(
            item["row"],
            item["score"],
            item["matched_fields"],
        )
        for item in scored
    ]

    results = unique_results_by_disease_id(raw_results, limit)

    return {
        "status": "success",
        "query": original_query,
        "normalized_query": query,
        "result_count": len(results),
        "results": results,
    }


# ============================================================
# GET SINGLE DISEASE
# ============================================================

@app.get("/api/v1/diseases/{disease_id}")
def get_disease(disease_id: int):
    """Return a clean, frontend-ready disease record."""
    if df is None or df.empty:
        raise HTTPException(
            status_code=500,
            detail="Knowledge base is not loaded",
        )

    matches = df[disease_id_matches(df["disease_id"], disease_id)]

    if matches.empty:
        raise HTTPException(
            status_code=404,
            detail="Disease not found",
        )

    row = matches.iloc[0]

    data = {
        "disease_id": first_available(row, ["disease_id"]),

        "basic_information": {
            "ayurvedic_name": first_available(row, ["ayurvedic_name"]),
            "disease_english": first_available(row, ["disease_english"]),
            "disease_hindi": first_available(row, ["disease_hindi"]),
            "disease_gujarati": first_available(row, ["disease_gujarati"]),
        },

        "symptoms": {
            "english": first_available(row, ["symptoms_english"]),
            "hindi": first_available(row, ["symptoms_hindi"]),
            "gujarati": first_available(row, ["symptoms_gujarati"]),
            "severity": first_available(row, ["symptom_severity"]),
        },

        "diagnosis": first_available(
            row,
            ["diagnosis_and_tests"],
        ),

        "treatment": {
            "english": first_available(row, ["treatment_english"]),
            "hindi": first_available(row, ["treatment_hindi"]),
            "gujarati": first_available(row, ["treatment_gujarati"]),
            "duration": first_available(row, ["duration_of_treatment"]),
        },

        "ayurvedic_medicine": {
            "english": first_available(
                row,
                ["ayurvedic_medicine_english", "medicine_english"],
            ),
            "hindi": first_available(
                row,
                ["ayurvedic_medicine_hindi", "medicine_hindi"],
            ),
            "gujarati": first_available(
                row,
                ["ayurvedic_medicine_gujarati", "medicine_gujarati"],
            ),
        },

        "herbs": first_available(
            row,
            ["Ayurvedic Herbs", "herbs_english", "herbs_hindi", "herbs_gujarati"],
        ),

        "formulation": first_available(
            row,
            [
                "Formulation",
                "formulation_english",
                "formulation_hindi",
                "formulation_gujarati",
            ],
        ),

        "doshas": {
            "original": first_available(row, ["doshas"]),
            "normalized": first_available(row, ["doshas_normalized"]),
        },

        "constitution": first_available(
            row,
            ["constitution_prakriti"],
        ),

        "diet_and_lifestyle": first_available(
            row,
            ["diet_and_lifestyle_recommendations"],
        ),

        "yoga_and_physical_therapy": first_available(
            row,
            ["yoga_and_physical_therapy"],
        ),

        "prevention": first_available(
            row,
            ["prevention"],
        ),

        "complications": first_available(
            row,
            ["complications"],
        ),

        "risk_factors": first_available(
            row,
            ["risk_factors"],
        ),

        "environmental_factors": first_available(
            row,
            ["environmental_factors"],
        ),

        "dietary_habits": first_available(
            row,
            ["dietary_habits"],
        ),

        "seasonal_variation": first_available(
            row,
            ["seasonal_variation"],
        ),

        "age_group": first_available(
            row,
            ["age_group"],
        ),

        "gender": first_available(
            row,
            ["gender"],
        ),

        "reference": {
            "text": first_available(
                row,
                ["reference_english", "reference_hindi", "reference_gujarati"],
            ),
            "source": first_available(
                row,
                ["reference_source"],
            ),
            "section": first_available(
                row,
                ["reference_section"],
            ),
        },
    }

    return {
        "status": "success",
        "data": data,
    }


# ============================================================
# GET RECOMMENDATION
# ============================================================

@app.get("/api/v1/recommend")
def recommend_get(
    q: str = Query(..., min_length=2),
):
    """
    Return the best matching Ayurvedic knowledge record for a query.
    Supports English, Hindi and Gujarati fields.
    """
    if df is None or df.empty:
        raise HTTPException(
            status_code=500,
            detail="Knowledge base is not loaded",
        )

    original_query = clean_value(q).strip()
    query = normalize_text(original_query)

    if not query:
        raise HTTPException(
            status_code=400,
            detail="Query cannot be empty",
        )

    query = SEARCH_ALIASES.get(query, query)

    # Stronger fields first.
    search_fields = [
        ("disease_english", 10),
        ("disease_hindi", 10),
        ("disease_gujarati", 10),
        ("disease_english_search", 10),
        ("disease_hindi_search", 10),
        ("disease_gujarati_search", 10),
        ("ayurvedic_name", 8),
        ("symptoms_english", 6),
        ("symptoms_hindi", 6),
        ("symptoms_gujarati", 6),
        ("treatment_english", 4),
        ("treatment_hindi", 4),
        ("treatment_gujarati", 4),
        ("ayurvedic_medicine_english", 3),
        ("ayurvedic_medicine_hindi", 3),
        ("ayurvedic_medicine_gujarati", 3),
        ("Ayurvedic Herbs", 2),
        ("Formulation", 2),
        ("doshas", 2),
        ("reference_english", 1),
        ("reference_hindi", 1),
        ("reference_gujarati", 1),
    ]

    matches = []

    for _, row in df.iterrows():
        score = 0
        matched_fields = []

        for column, field_score in search_fields:
            if column not in df.columns:
                continue

            value = normalize_text(row[column])

            if value and query in value:
                score += field_score
                matched_fields.append(column)

        if score > 0:
            matches.append({
                "score": score,
                "matched_fields": matched_fields,
                "row": row,
            })

    if not matches:
        return {
            "status": "not_found",
            "query": original_query,
            "message": "No matching Ayurvedic knowledge found.",
        }

    matches.sort(
        key=lambda item: item["score"],
        reverse=True,
    )

    best = matches[0]["row"]

    recommendation = {
        "disease_id": first_available(best, ["disease_id"]),
        "disease": first_available(best, ["disease_english"]),
        "disease_hindi": first_available(best, ["disease_hindi"]),
        "disease_gujarati": first_available(best, ["disease_gujarati"]),
        "ayurvedic_name": first_available(best, ["ayurvedic_name"]),

        "symptoms": first_available(
            best,
            ["symptoms_english", "symptoms_hindi", "symptoms_gujarati"],
        ),

        "treatment": first_available(
            best,
            ["treatment_english", "treatment_hindi", "treatment_gujarati"],
        ),

        "ayurvedic_medicine": first_available(
            best,
            [
                "ayurvedic_medicine_english",
                "medicine_english",
                "ayurvedic_medicine_hindi",
                "medicine_hindi",
                "ayurvedic_medicine_gujarati",
                "medicine_gujarati",
            ],
        ),

        "herbs": first_available(
            best,
            ["Ayurvedic Herbs", "herbs_english", "herbs_hindi", "herbs_gujarati"],
        ),

        "formulation": first_available(
            best,
            [
                "Formulation",
                "formulation_english",
                "formulation_hindi",
                "formulation_gujarati",
            ],
        ),

        "doshas": first_available(
            best,
            ["doshas", "doshas_normalized"],
        ),

        "diet_and_lifestyle": first_available(
            best,
            ["diet_and_lifestyle_recommendations"],
        ),

        "yoga_and_physical_therapy": first_available(
            best,
            ["yoga_and_physical_therapy"],
        ),

        "prevention": first_available(
            best,
            ["prevention"],
        ),

        "complications": first_available(
            best,
            ["complications"],
        ),

        "reference": first_available(
            best,
            ["reference_english", "reference_hindi", "reference_gujarati"],
        ),

        "reference_source": first_available(
            best,
            ["reference_source"],
        ),

        "reference_section": first_available(
            best,
            ["reference_section"],
        ),
    }

    return {
        "status": "success",
        "query": original_query,
        "normalized_query": query,
        "match_score": matches[0]["score"],
        "matched_fields": matches[0]["matched_fields"],
        "recommendation": recommendation,
        "disclaimer": (
            "Educational and knowledge-retrieval output only. "
            "It is not a substitute for diagnosis or professional medical advice."
        ),
    }


# ============================================================
# POST RECOMMENDATION REQUEST MODEL
# ============================================================

class RecommendationRequest(BaseModel):
    symptoms: Optional[str] = None
    disease: Optional[str] = None
    dosha: Optional[str] = None
    property: Optional[str] = None


# ============================================================
# POST RECOMMENDATION ENGINE
# ============================================================

@app.post("/api/v1/recommend")
def recommend_post(request: RecommendationRequest):
    """
    Score records against symptoms, disease, dosha and property.

    This is a deterministic knowledge-base ranking engine, not a
    medical diagnosis model.
    """
    if df is None or df.empty:
        raise HTTPException(
            status_code=500,
            detail="Knowledge base is not loaded",
        )

    if not any([
        request.symptoms,
        request.disease,
        request.dosha,
        request.property,
    ]):
        raise HTTPException(
            status_code=400,
            detail=(
                "Provide at least one of symptoms, disease, "
                "dosha or property."
            ),
        )

    symptoms = normalize_text(request.symptoms)
    disease = normalize_text(request.disease)
    dosha = normalize_text(request.dosha)
    prop = normalize_text(request.property)

    # Apply the same Gujarati alias logic used by GET search.
    disease = SEARCH_ALIASES.get(disease, disease)
    symptoms = SEARCH_ALIASES.get(symptoms, symptoms)
    dosha = SEARCH_ALIASES.get(dosha, dosha)
    prop = SEARCH_ALIASES.get(prop, prop)

    scored = []

    for _, row in df.iterrows():
        score = 0
        matched = []

        # --------------------------------------------------------
        # Disease matching
        # --------------------------------------------------------
        if disease:
            disease_fields = [
                "disease_english",
                "disease_hindi",
                "disease_gujarati",
                "disease_english_search",
                "disease_hindi_search",
                "disease_gujarati_search",
                "ayurvedic_name",
            ]

            if any(
                disease in normalize_text(row[column])
                for column in disease_fields
                if column in df.columns
            ):
                score += 5
                matched.append("disease")

        # --------------------------------------------------------
        # Symptom matching
        # --------------------------------------------------------
        if symptoms:
            symptom_text = " ".join(
                normalize_text(row[column])
                for column in [
                    "symptoms_english",
                    "symptoms_hindi",
                    "symptoms_gujarati",
                    "symptoms_search",
                ]
                if column in df.columns
            )

            symptom_words = [
                word.strip()
                for word in symptoms.split(",")
                if word.strip()
            ]

            symptom_matches = 0

            for word in symptom_words:
                if word in symptom_text:
                    symptom_matches += 1

            if symptom_matches:
                score += 3 * symptom_matches
                matched.append(
                    f"symptoms:{symptom_matches}"
                )

        # --------------------------------------------------------
        # Dosha matching
        # --------------------------------------------------------
        if dosha:
            dosha_text = " ".join(
                normalize_text(row[column])
                for column in [
                    "doshas",
                    "doshas_normalized",
                    "doshas_english",
                    "doshas_hindi",
                    "doshas_gujarati",
                    "doshas_search",
                ]
                if column in df.columns
            )

            if dosha in dosha_text:
                score += 2
                matched.append("dosha")

        # --------------------------------------------------------
        # Property / treatment / herb / formulation matching
        # --------------------------------------------------------
        if prop:
            property_text = " ".join(
                normalize_text(row[column])
                for column in [
                    "treatment_english",
                    "treatment_hindi",
                    "treatment_gujarati",
                    "Ayurvedic Herbs",
                    "herbs_english",
                    "herbs_hindi",
                    "herbs_gujarati",
                    "Formulation",
                    "formulation_english",
                    "formulation_hindi",
                    "formulation_gujarati",
                    "diet_and_lifestyle_recommendations",
                    "medicine_english",
                    "medicine_hindi",
                    "medicine_gujarati",
                    "ayurvedic_medicine_english",
                    "ayurvedic_medicine_hindi",
                    "ayurvedic_medicine_gujarati",
                ]
                if column in df.columns
            )

            if prop in property_text:
                score += 2
                matched.append("property")

        if score > 0:
            scored.append({
                "score": score,
                "matched": matched,
                "row": row,
            })

    scored.sort(
        key=lambda item: item["score"],
        reverse=True,
    )

    raw_results = []

    for item in scored:
        row = item["row"]

        raw_results.append({
            "disease_id": first_available(row, ["disease_id"]),
            "ayurvedic_name": first_available(row, ["ayurvedic_name"]),
            "disease": first_available(row, ["disease_english"]),
            "disease_hindi": first_available(row, ["disease_hindi"]),
            "disease_gujarati": first_available(row, ["disease_gujarati"]),
            "matched": item["matched"],
            "score": item["score"],
            "symptoms": first_available(
                row,
                ["symptoms_english", "symptoms_hindi", "symptoms_gujarati"],
            ),
            "medicines": first_available(
                row,
                [
                    "ayurvedic_medicine_english",
                    "medicine_english",
                    "ayurvedic_medicine_hindi",
                    "medicine_hindi",
                    "ayurvedic_medicine_gujarati",
                    "medicine_gujarati",
                ],
            ),
            "herbs": first_available(
                row,
                ["Ayurvedic Herbs", "herbs_english", "herbs_hindi", "herbs_gujarati"],
            ),
            "formulation": first_available(
                row,
                [
                    "Formulation",
                    "formulation_english",
                    "formulation_hindi",
                    "formulation_gujarati",
                ],
            ),
            "doshas": first_available(
                row,
                ["doshas", "doshas_normalized"],
            ),
            "reference": first_available(
                row,
                ["reference_english", "reference_hindi", "reference_gujarati"],
            ),
        })

    results = unique_results_by_disease_id(raw_results, 10)

    return {
        "status": "success",
        "query": request.model_dump(),
        "result_count": len(results),
        "recommendations": results,
        "disclaimer": (
            "Educational and knowledge-retrieval output only. "
            "It is not a substitute for diagnosis or professional medical advice."
        ),
    }


# ============================================================
# DEVELOPMENT ENTRY POINT
# ============================================================

if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "main:app",
        host="127.0.0.1",
        port=8000,
        reload=True,
    )
