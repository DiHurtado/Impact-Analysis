import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Impact Analysis Tool", layout="wide")

st.title("🚨 Impact Analysis Tool")

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
# LOAD POLARION
# ===============================
def load_polarion():
    path = r"C:\Users\dihurtado\BorgWarner\17001357 AM_Stellantis _PE INV_150kW and 250kW PIM_MY24 - 05-Metrics\PolarionInternalWorkItemsSTLA_400V.xlsm"

    try:
        df = pd.read_excel(path, engine="openpyxl", sheet_name="Results", skiprows=4)
        df.columns = df.columns.astype(str).str.strip().str.lower()
        return df
    except Exception as e:
        st.warning(f"⚠️ Polarion file error: {e}")
        return None


# ===============================
# EXTRAER ID
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

    jira_test_col = "test_case_id" if "test_case_id" in df.columns else None

    if jira_test_col is None:
        df["polarion_test_match"] = "NO"
        df["polarion_match"] = "NO"
        df["safety_flag"] = ""
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
        raw = row.get(jira_test_col)
        test_id = extract_test_case_id(raw)

        if test_id in pol_lookup:
            return pd.Series(["YES", pol_lookup[test_id]])

        return pd.Series(["NO", ""])

    df[["polarion_test_match", "safety_flag"]] = df.apply(process_row, axis=1)
    df["polarion_match"] = df["polarion_test_match"]

    return df


# ===============================
# READ JIRA
# ===============================
def read_jira_file(file):
    try:
        tables = pd.read_html(file)
        return tables[1] if len(tables) > 1 else None
    except:
        return None


# ===============================
# NORMALIZE
# ===============================
def normalize_columns(df):
    cols = []
    for col in df.columns:
        c = str(col).strip().lower()
        cols.append(COLUMN_MAPPING.get(c, c))
    df.columns = cols
    return df


# ===============================
# FIX DUPLICATES
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
# ✅ SCORING + NATURAL JUSTIFICATION
# ===============================
def calculate_new_score(row):

    # PRIORITY
    pr_map = {"blocker": 4, "high": 3, "medium": 2, "low": 1}
    priority_raw = str(row.get("priority", "")).lower()
    priority_val = pr_map.get(priority_raw, 1)

    priority_text = {
        "blocker": "blocker",
        "high": "alta",
        "medium": "media",
        "low": "baja"
    }.get(priority_raw, "baja")

    # ASIL
    safety = str(row.get("safety_flag", "")).lower()
    asil_map = {"asil d": 5, "asil c": 4, "asil b": 3, "asil a": 2, "qm": 1}

    asil_val = 1
    asil_label = "QM"

    for k, v in asil_map.items():
        if k in safety:
            asil_val = v
            asil_label = k.upper()
            break

    # FOUND IN
    found = str(row.get("found_in", "")).lower()

    fi_map = {
        "by customer": (5, "By Customer"),
        "system qualification test": (4, "System Qualification Test"),
        "system integration test": (3, "System Integration Test"),
        "software qualification test": (4, "Software Qualification Test"),
        "software integration test": (3, "Software Integration Test"),
        "software unit test": (2, "Software Unit Test")
    }

    fi_val = 1
    fi_label = "Other"

    for k, (v, label) in fi_map.items():
        if k in found:
            fi_val = v
            fi_label = label
            break

    # PRELIMINAR
    prelim = priority_val * fi_val

    if prelim >= 10:
        prelim_severity = "High"
        score = 3
    elif prelim <= 3:
        prelim_severity = "Low"
        score = 1
    else:
        prelim_severity = "Medium"
        score = 2

    # FINAL
    final_calc = score * asil_val

    if final_calc >= 10:
        final_severity = "High"
        action = "Immediate cross-functional action required"
    elif final_calc <= 3:
        final_severity = "Low"
        action = "Monitor"
    else:
        final_severity = "Medium"
        action = "Plan fix in next release"

    # ✅ JUSTIFICACIÓN NATURAL
    justification = (
        f"La prioridad del defecto es {priority_text}, "
        f"el nivel de seguridad es {asil_label} "
        f"y fue encontrado en {fi_label}. "
        f"Con base en estos factores, el impacto final se clasifica como {final_severity} "
        f"y se recomienda la siguiente acción: {action}."
    )

    return pd.Series([
        priority_val,
        fi_val,
        asil_val,
        prelim,
        prelim_severity,
        score,
        final_calc,
        final_severity,
        action,
        justification
    ])


# ===============================
# APP
# ===============================
uploaded_file = st.file_uploader(
    "Upload JIRA Excel (.xls)",
    type=["xls"],
    accept_multiple_files=False
)

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

    df_calc[[
        "Priority_Value",
        "FoundIn_Value",
        "ASIL_Value",
        "Preliminary_Result",
        "Preliminary_Severity",
        "Score",
        "Final_Result",
        "Impact Level",
        "Recommended Action",
        "Justification"
    ]] = df_calc.apply(calculate_new_score, axis=1)

    severity_order = {"High": 0, "Medium": 1, "Low": 2}
    df_calc["severity_order"] = df_calc["Impact Level"].map(severity_order)

    df_calc = df_calc.sort_values(by=["severity_order", "Final_Result"], ascending=[True, False])

    columns_to_hide = [
        "polarion_test_match",
        "polarion_match",
        "safety_flag",
        "Priority_Value",
        "FoundIn_Value",
        "ASIL_Value",
        "Preliminary_Severity",
        "Preliminary_Result",
        "Score",
        "Final_Result",
        "severity_order"
    ]

    df_display = df_calc.drop(columns=[c for c in columns_to_hide if c in df_calc.columns])

    st.subheader("🚨 Impact Analysis (ASPICE Ready)")
    st.dataframe(df_display, use_container_width=True)

    st.subheader("🔥 High Impact Defects")
    st.dataframe(df_display[df_display["Impact Level"] == "High"], use_container_width=True)

    buffer = io.BytesIO()
    df_display.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)

    st.download_button(
        label="📥 Download Analysis",
        data=buffer,
        file_name="impact_analysis_clean.xlsx"
    )
