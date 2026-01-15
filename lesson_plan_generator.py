import streamlit as st
from docx import Document
import io
import json
import os

st.set_page_config(page_title="Plan Filler: Weekly Lesson Plan", layout="wide")
st.title("Plan Filler: Weekly Lesson Plan Generator")

st.markdown("""
Upload your Word template or use the default WLPT template, then paste the JSON of your weekly lesson plan.
Click **Generate Lesson Plan** to get a ready-to-use Word document.
""")

# Default template path
default_template_path = os.path.join("templates", "WLPT.docx")
template_file = st.file_uploader("Upload your Lesson Plan Template (.docx)", type=["docx"])
if not template_file:
    if os.path.exists(default_template_path):
        template_file = default_template_path
    else:
        st.warning("No template uploaded or found in repo.")
        st.stop()

# JSON input
json_input = st.text_area("Paste your weekly lesson plan JSON here:")

if st.button("Generate Lesson Plan"):
    if not json_input.strip():
        st.error("Please paste your JSON data.")
    else:
        try:
            lesson_data = json.loads(json_input)
            doc = Document(template_file)

            # Replace placeholders in paragraphs
            for paragraph in doc.paragraphs:
                for key, value in lesson_data.items():
                    placeholder = f"{{{{{key}}}}}"
                    if placeholder in paragraph.text:
                        paragraph.text = paragraph.text.replace(placeholder, str(value))

            # Replace placeholders in tables
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for key, value in lesson_data.items():
                            placeholder = f"{{{{{key}}}}}"
                            if placeholder in cell.text:
                                cell.text = cell.text.replace(placeholder, str(value))

            # Save to in-memory file
            output = io.BytesIO()
            doc.save(output)
            output.seek(0)

            # Download button
            st.download_button(
                label="Download Completed Lesson Plan",
                data=output,
                file_name="Weekly_Lesson_Plan.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except json.JSONDecodeError:
            st.error("Invalid JSON. Please check your data format.")
        except Exception as e:
            st.error(f"An error occurred: {e}")
