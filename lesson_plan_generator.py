# PLAN FILLER – WEEKLY LESSON PLAN GENERATOR
# JSON → WORD (TEMPLATE UPLOAD SUPPORTED)

import streamlit as st
import json
import io
from docx import Document

st.set_page_config(page_title="Plan Filler", layout="wide")

st.title("📘 Plan Filler – Weekly Lesson Plan Generator")

st.markdown(
    "1️⃣ Upload your **Word lesson plan template**  \n"
    "2️⃣ Paste the **lesson plan JSON**  \n"
    "3️⃣ Download the populated document"
)

# ==================================================
# CORRECT + FINAL JSON SKELETON (MATCHES TEMPLATE)
# ==================================================

JSON_SKELETON = {
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

# ==================================================
# TEMPLATE UPLOAD
# ==================================================

uploaded_file = st.file_uploader(
    "📤 Upload Word Template (.docx)",
    type=["docx"]
)

# ==================================================
# JSON INPUT
# ==================================================

json_input = st.text_area(
    "📄 Lesson Plan JSON",
    value=json.dumps(JSON_SKELETON, indent=2),
    height=550
)

# ==================================================
# HELPER FUNCTIONS
# ==================================================

def flatten_json(data):
    flat = {}
    for key, value in data.items():
        if isinstance(value, dict):
            for subkey, subvalue in value.items():
                flat[f"{{{{{subkey}}}}}"] = str(subvalue)
        else:
            flat[f"{{{{{key}}}}}"] = str(value)
    return flat


def replace_placeholders(doc, replacements):
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in replacements.items():
                        if key in paragraph.text:
                            paragraph.text = paragraph.text.replace(key, value)

# ==================================================
# GENERATE DOCUMENT
# ==================================================

if st.button("🚀 Generate Lesson Plan", use_container_width=True):
    if not uploaded_file:
        st.error("Please upload a Word template.")
    else:
        try:
            data = json.loads(json_input)
            replacements = flatten_json(data)

            doc = Document(uploaded_file)
            replace_placeholders(doc, replacements)

            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)

            filename = f"{data.get('Teacher','Teacher')}_{data.get('WeekNumber','Week')}_Lesson_Plan.docx"

            st.success("✅ Lesson plan generated successfully.")

            st.download_button(
                "📥 Download Word File",
                data=buffer,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

        except json.JSONDecodeError:
            st.error("Invalid JSON format.")
        except Exception as e:
            st.error(str(e))
