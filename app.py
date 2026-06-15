import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Impact Analysis Tool", layout="wide")

st.title("🚨 Impact Analysis Tool")

# ===============================
# CONFIG
# ===============================
REQUIRED_COLUMNS = [
    "priority",
    "found_in",
    "test_case_id"
]

COLUMN_MAPPING = {
    "test case id": "test_case_id",
    "test_case_id": "test_case_id",
    "priority": "priority",
    "found in": "found_in"
}

# ===============================
# ✅ LOAD POLARION DESDE GITHUB
# ===============================
@st.cache_data
def load_polarion():
    try:
        df = pd.read_excel(
            "data/PolarionInternalWorkItemsSTLA_400V.xlsm",
            engine="openpyxl",
            sheet_name="Results",
            skiprows=4
        )

        df.columns = df.columns.astype(str).str.strip().str.lower()

        return df

    except Exception as e:
        st.warning(f"⚠️ Polarion file error: {e}")
        return None


# ===============================
# EXTRAER TEST ID
# ===============================
def extract_test_case_id(value):
    if pd.isna(value):
        return None
    value = str(value)
    if "=" in value:
        return value.split("=")[-1].strip()
    return value.strip()


# ===============================
# MERGE POLARION
# ===============================
def merge_polarion(df, df_pol):

    if df_pol is None:
        df["polarion_test_match"] = "NO"
        df["polarion_match"] = "NO"
        df["safety_flag"] = ""
        return df

    if "test_case_id" not in df.columns:
        return df

    if "verification case id" not in df_pol.columns:
        st.error(f"❌ Columns available: {df_pol.columns.tolist()}")
        return df

    if "safety" not in df_pol.columns:
        df_pol["safety"] = ""

    pol_lookup = {
        str(v).strip(): str(s).strip()
        for v, s in zip(df_pol["verification case id"], df_pol["safety"])
        if pd.notna(v)
    }

    def process_row(row):
        test_id = extract_test_case_id(row.get("test_case_id"))

        if test_id in pol_lookup:
            return pd.Series(["YES", pol_lookup[test_id]])

        return pd.Series(["NO", ""])

    df[["polarion_test_match", "safety_flag"]] = df.apply(process_row, axis=1)
    df["polarion_match"] = df["polarion_test_match"]

    return df


# ===============================
# LEER JIRA
# ===============================
def read_jira_file(file):
    try:
        tables = pd.read_html(file)
        return tables[1] if len(tables) > 1 else None
    except:
        return None


# ===============================
# NORMALIZAR
# ===============================
def normalize_columns(df):
    df.columns = [
        COLUMN_MAPPING.get(str(c).strip().lower(), str(c).strip().lower())
        for c in df.columns
    ]
    return df


# ===============================
# SCORING + JUSTIFICACIÓN
# ===============================
def calculate_score(row):

    pr_map = {"blocker": 4, "high": 3, "medium": 2, "low": 1}
    priority = str(row.get("priority", "")).lower()
    pr_val = pr_map.get(priority, 1)

    priority_text = {
        "blocker": "crítica",
        "high": "alta",
        "medium": "media",
        "low": "baja"
    }.get(priority, "baja")

    safety = str(row.get("safety_flag", "")).lower()
    asil_map = {"asil d": 5, "asil c": 4, "asil b": 3, "asil a": 2, "qm": 1}

    asil_val = 1
    asil_label = "QM"

    for k, v in asil_map.items():
        if k in safety:
            asil_val = v
            asil_label = k.upper()
            break

    fi_map = {
        "by customer": (5, "By Customer"),
        "system qualification": (4, "System Qualification Test"),
        "integration": (3, "System Integration Test"),
        "software qualification": (4, "Software Qualification Test")
    }

    fi_val = 1
    fi_label = "Other"

    found = str(row.get("found_in", "")).lower()
    for k, (v, label) in fi_map.items():
        if k in found:
            fi_val = v
            fi_label = label
            break

    prelim = pr_val * fi_val

    if prelim >= 10:
        sev = "High"
    elif prelim <= 3:
        sev = "Low"
    else:
        sev = "Medium"

    final_score = asil_val * (3 if sev == "High" else 2 if sev == "Medium" else 1)

    if final_score >= 10:
        impact = "High"
        action = "Immediate cross-functional action required"
    elif final_score <= 3:
        impact = "Low"
        action = "Monitor"
    else:
        impact = "Medium"
        action = "Plan fix in next release"

    justification = (
        f"La prioridad del defecto es {priority_text}, "
        f"el ASIL es {asil_label} y fue encontrado en {fi_label}. "
        f"Con base en esto, el impacto es {impact} y se recomienda: {action}."
    )

    return pd.Series([impact, action, justification])


# ===============================
# UI
# ===============================
uploaded_file = st.file_uploader(
    "Upload JIRA Excel (.xls)",
    type=["xls"],
    accept_multiple_files=False
)

if uploaded_file:

    df = read_jira_file(uploaded_file)

    if df is None:
        st.error("❌ Could not read JIRA file")
        st.stop()

    df = normalize_columns(df)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]

    if missing:
        st.error(f"❌ Missing columns: {missing}")
        st.stop()

    df_pol = load_polarion()
    df = merge_polarion(df, df_pol)

    df[["Impact Level", "Recommended Action", "Justification"]] = df.apply(
        calculate_score,
        axis=1
    )

    st.subheader("🚨 Impact Analysis")
    st.dataframe(df, use_container_width=True)

    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)

    st.download_button(
        "📥 Download Results",
        buffer,
        file_name="impact_analysis.xlsx"
    )
``
