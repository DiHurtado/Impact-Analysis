import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="PRM Defect Analysis", layout="wide")

st.title("🚨 Problem Resolution Management Tool")

REQUIRED_COLUMNS = [
    "description",
    "responsible_competency",
    "priority",
    "found_in"
]

COLUMN_MAPPING = {
    "description": "description",
    "priority": "priority",
    "responsible competency": "responsible_competency",
    "found in": "found_in"
}

# ===============================
# ✅ LOAD POLARION
# ===============================
def load_polarion():
    path = r"C:\Users\dihurtado\BorgWarner\17001357 AM_Stellantis _PE INV_150kW and 250kW PIM_MY24 - 05-Metrics\PolarionInternalWorkItemsSTLA_400V.xlsm"

    try:
        df = pd.read_excel(
            path,
            engine="openpyxl",
            sheet_name="Results",
            skiprows=4
        )

        df.columns = (
            df.columns.astype(str)
            .str.strip()
            .str.lower()
        )

        return df

    except Exception as e:
        st.warning(f"⚠️ Polarion file error: {e}")
        return None


# ===============================
# ✅ EXTRAER ID DESDE URL
# ===============================
def extract_test_case_id(value):

    if pd.isna(value):
        return None

    value = str(value)

    if "=" in value:
        return value.split("=")[-1].strip()

    return value.strip()


# ===============================
# ✅ MERGE CON POLARION + SAFETY
# ===============================
def merge_polarion(df, df_pol):

    if df_pol is None:
        df["polarion_test_match"] = "NO"
        df["polarion_match"] = "NO"
        df["safety_flag"] = ""
        return df

    # detectar columna en Jira
    jira_test_col = None
    for c in df.columns:
        if "test case id" in c.lower():
            jira_test_col = c
            break

    if jira_test_col is None:
        df["polarion_test_match"] = "NO"
        df["polarion_match"] = "NO"
        df["safety_flag"] = ""
        return df

    # columnas Polarion
    if "verification case id" not in df_pol.columns:
        st.error(f"❌ Columns available: {df_pol.columns.tolist()}")
        return df

    if "safety" not in df_pol.columns:
        df_pol["safety"] = ""

    # ✅ crear lookup ID → Safety
    pol_lookup = {
        str(v).strip(): str(s).strip()
        for v, s in zip(df_pol["verification case id"], df_pol["safety"])
        if pd.notna(v)
    }

    def process_row(row):

        raw = row.get(jira_test_col)
        test_id = extract_test_case_id(raw)

        if test_id in pol_lookup:
            return pd.Series(["YES", pol_lookup[test_id]])

        return pd.Series(["NO", ""])

    df[["polarion_test_match", "safety_flag"]] = df.apply(process_row, axis=1)

    df["polarion_match"] = df["polarion_test_match"]

    return df


# ===============================
# ✅ READ JIRA
# ===============================
def read_jira_file(file):
    try:
        tables = pd.read_html(file)
        return tables[1] if len(tables) > 1 else None
    except:
        return None


# ===============================
# ✅ CLEAN COLUMNS
# ===============================
def normalize_columns(df):
    cols = []
    for col in df.columns:
        c = str(col).strip().lower()
        cols.append(COLUMN_MAPPING.get(c, c))
    df.columns = cols
    return df


# ===============================
# ✅ FIX DUPLICATES
# ===============================
def fix_duplicates(df):
    seen = {}
    new_cols = []

    for col in df.columns:
        if col not in seen:
            seen[col] = 0
            new_cols.append(col)
        else:
            seen[col] += 1
            new_cols.append(f"{col}_{seen[col]}")

    df.columns = new_cols
    return df


# ===============================
# ✅ PRM ENGINE + SAFETY
# ===============================
def calculate_prm(row):

    score = 0

    priority = str(row.get("priority", "")).lower()
    found_in = str(row.get("found_in", "")).lower()
    competency = str(row.get("responsible_competency", "")).lower()
    description = str(row.get("description", "")).lower()
    trace = str(row.get("polarion_match", "")).lower()
    safety = str(row.get("safety_flag", "")).lower()

    # ✅ BASE SCORING
    if priority == "high":
        score += 40
    elif priority == "medium":
        score += 25
    elif priority == "low":
        score += 10

    if "customer" in found_in:
        score += 25
    elif "system" in found_in:
        score += 15
    elif "integration" in found_in:
        score += 10

    if competency == "system":
        score += 15
    elif competency == "hardware":
        score += 12
    elif competency in ["sw", "software"]:
        score += 10

    # ✅ DESCRIPCIÓN
    if any(x in description for x in ["failure", "shutdown", "crash", "fault", "voltage", "battery"]):
        score += 25

    # ✅ TRACEABILITY
    if trace == "no":
        score += 15

    # ✅ 🔥 SAFETY (NUEVO)
    if "asil d" in safety:
        score += 30
    elif "asil c" in safety:
        score += 25
    elif "asil b" in safety:
        score += 20
    elif "asil a" in safety:
        score += 10
    elif "yes" in safety:
        score += 20

    # ✅ IMPACT
    if score >= 70:
        impact = "High"
    elif score >= 40:
        impact = "Medium"
    else:
        impact = "Low"

    # ✅ PRM PRIORITY
    if "customer" in found_in and priority == "high":
        prm_priority = "CRITICAL"
        action = "Immediate escalation. Customer impact."
    elif score >= 80:
        prm_priority = "CRITICAL"
        action = "Immediate cross-functional action required."
    elif score >= 60:
        prm_priority = "HIGH"
        action = "Fix before next release."
    elif score >= 40:
        prm_priority = "MEDIUM"
        action = "Track and plan resolution."
    else:
        prm_priority = "LOW"
        action = "Monitor. No immediate action."

    justification = f"Score={score}, priority={priority}, found_in={found_in}, trace={trace}, safety={safety}"

    return pd.Series([score, impact, prm_priority, action, justification])


# ===============================
# 🚀 APP
# ===============================
uploaded_file = st.file_uploader("Upload JIRA Excel (.xls)", type=["xls"])

if uploaded_file:

    df = read_jira_file(uploaded_file)

    if df is None:
        st.error("❌ Could not extract table")
        st.stop()

    df = normalize_columns(df)
    df = fix_duplicates(df)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]

    if missing:
        st.error(f"❌ Missing columns: {missing}")
        st.stop()

    df_pol = load_polarion()
    df = merge_polarion(df, df_pol)

    df_calc = df.copy()

    df_calc[["Score", "Impact", "PRM Priority",
             "Recommended Action", "Justification"]] = df_calc.apply(calculate_prm, axis=1)

    priority_order = {"CRITICAL": 0, "HIGH": 1, "MEDIUM": 2, "LOW": 3}
    df_calc["priority_order"] = df_calc["PRM Priority"].map(priority_order)

    df_calc = df_calc.sort_values(by=["priority_order", "Score"], ascending=[True, False])

    st.subheader("🚨 Defect Resolution Priority")
    st.dataframe(df_calc.drop(columns=["priority_order"]), use_container_width=True)

    st.subheader("🔥 Critical Defects (Immediate Attention)")
    st.dataframe(df_calc[df_calc["PRM Priority"] == "CRITICAL"], use_container_width=True)

    # ✅ Export
    buffer = io.BytesIO()
    df_calc.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)

    st.download_button(
        label="📥 Download PRM Analysis",
        data=buffer,
        file_name="prm_analysis.xlsx"
    )