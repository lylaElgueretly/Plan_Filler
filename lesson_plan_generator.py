# WEEKLY LESSON PLAN GENERATOR
# JSON → WORD (WITH TEMPLATE UPLOAD)
# PLACEHOLDERS MATCH WORD TEMPLATE EXACTLY

import streamlit as st
import json
import tempfile
import os
from docx import Document

st.set_page_config(page_title="Weekly Lesson Plan Generator", layout="wide")

st.title("📘 Weekly Lesson Plan Generator")

st.write(
    "1️⃣ Upload the **Word lesson plan template**  
     2️⃣ Paste the **lesson plan JSON**  
     3️⃣ Download the populated document"
)

# =========================
# CORRECT JSON SKELETON
# =========================

CORRECT_JSON_TEMPLATE = {
    "Teacher": "",
    "WeekNumber": "",
    "YearClass": "",
    "Subject": "",
    "UnitTopic": "",

    "LearningObjective": {
        "Class1_LearningObjective": "",
        "Class2_LearningObjective": "",
        "Class3_LearningObjective": "",
        "Class4_LearningObjective": "",
        "Class5_LearningObjective": ""
    },

    "SuccessCriteria": {
        "Class1_SuccessCriteria": "",
        "Class2_SuccessCriteria": "",
        "Class3_SuccessCriteria": "",
        "Class4_SuccessCriteria": "",
        "Class5_SuccessCriteria": ""
    },

    "KeyVocabulary": {
        "Class1_Vocabulary": "",
        "Class2_Vocabulary": "",
        "Class3_Vocabulary": "",
        "Class4_Vocabulary": "",
        "Class5_Vocabulary": ""
    },

    "KeyQuestions": {
        "Class1_KeyQuestions": "",
        "Class2_KeyQuestions": "",
        "Class3_KeyQuestions": "",
        "Class4_KeyQuestions": "",
        "Class5_KeyQuestions": ""
    },

    "StarterActivity": {
        "Class1_StarterActivity": "",
        "Class2_StarterActivity": "",
        "Class3_StarterActivity": "",
        "Class4_StarterActivity": "",
        "Class5_StarterActivity": ""
    },

    "MainTeaching": {
        "Class1_MainTeaching": "",
        "Class2_MainTeaching": "",
        "Class3_MainTeaching": "",
        "Class4_MainTeaching": "",
        "Class5_MainTeaching": ""
    },

    "DifferentiatedActivities": {
        "Class1_DifferentiatedActivities": "",
        "Class2_DifferentiatedActivities": "",
        "Class3_DifferentiatedActivities": "",
        "Class4_DifferentiatedActivities": "",
        "Class5_DifferentiatedActivities": ""
    },

    "Plenary": {
        "Class1_Plenary": "",
        "Class2_Plenary": "",
        "Class3_Plenary": "",
        "Class4_Plenary": "",
        "Class5_Plenary": ""
    },

    "Reflection": {
        "Class1_Reflection": "",
        "Class2_Reflection": "",
        "Class3_Reflection": "",
        "Class4_Reflection": "",
        "Class5_Reflection": ""
    },

    "Homework": {
        "Class1_Homework": "",
        "Class2_Homework": "",
        "Class3_Homework": "",
        "Class4_Homework": "",
        "Class5_Homework": ""
    }
}

# =========================
# TEMPLATE UPLOAD
# =========================

uploaded_template = st.file_uploader(
    "📤 Upload Word Lesson Plan Template (.docx)",
    type=["docx"]
)

# =========================
# JSON INPUT
# =========================

json_input = st.text_area(
    "📄 Lesson Plan JSON (do NOT rename keys)",
    value=json.dumps(CORRECT_JSON_TEMPLATE, indent=2),
    height=600
)

# =========================
# HELPERS
# =========================

def flatten_json(data):
    flat = {}
    for key, value in data.items():
        if isinstance(value, dict):
            for sub_key, sub_value in value.items():
                flat[sub_key] = str(sub_value)
        else:
            flat[key] = str(value)
    return flat


def replace_placeholders(document, replacements):
    # Paragraphs
    for paragraph in document.paragraphs:
        for key, value in replacements.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, value)

    # Tables
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in replacements.items():
                    placeholder = f"{{{{{key}}}}}"
                    if placeholder in cell.text:
                        cell.text = cell.text.replace(placeholder, value)

# =========================
# GENERATE DOCUMENT
# =========================

if st.button("📄 Generate Lesson Plan"):
    if not uploaded_template:
        st.error("Please upload the Word template.")
        st.stop()

    try:
        lesson_data = json.loads(json_input)
    except json.JSONDecodeError as e:
        st.error(f"Invalid JSON: {e}")
        st.stop()

    # Save uploaded template temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(uploaded_template.read())
        template_path = tmp.name

    document = Document(template_path)
    flat_data = flatten_json(lesson_data)

    replace_placeholders(document, flat_data)

    # Save output
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as output:
        document.save(output.name)
        output_path = output.name

    with open(output_path, "rb") as f:
        st.download_button(
            "⬇ Download Completed Lesson Plan",
            data=f,
            file_name="Weekly_Lesson_Plan.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    st.success("Lesson plan generated successfully.")
