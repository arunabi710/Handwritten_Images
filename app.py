import streamlit as st
import io
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
from docx import Document
import pandas as pd

# Azure Form Recognizer credentials
endpoint = "https://arun-document-ai.cognitiveservices.azure.com/"
credential = AzureKeyCredential("d3abb1fb970e41d8b7f3330e202f342a")
document_analysis_client = DocumentAnalysisClient(endpoint, credential)
model_id = "testing-sow"

def process_pdf(file):
    # Process the uploaded PDF
    document = file.read()
    
    poller = document_analysis_client.begin_analyze_document(model_id, document)
    result = poller.result()

    # Initialize lists to store data
    project_scope = [result.documents[0].fields.get("Project Scope", {}).value]
    period_of_performance = [result.documents[0].fields.get("Period of Performance", {}).value]
    total_project_price = [result.documents[0].fields.get("Total Project Price", {}).value]
    deliverables = []

    # Handle Deliverables field which may have multiple items
    deliverables_field = result.documents[0].fields.get("Deliverables", {}).value
    if isinstance(deliverables_field, list):
        for deliverable in deliverables_field:
            if hasattr(deliverable, 'value') and isinstance(deliverable.value, dict):
                deliverable_value = deliverable.value.get('Deliverables', {}).value
                if deliverable_value:
                    deliverables.append(deliverable_value)
            else:
                st.warning(f"Unexpected data structure for deliverable: {deliverable}")
    else:
        st.warning("Deliverables data is not in expected list format.")

    # Determine the maximum length among all lists
    max_length = max(len(project_scope), len(period_of_performance), len(total_project_price), len(deliverables))

    # Fill missing data with None to align lengths
    project_scope += [None] * (max_length - len(project_scope))
    period_of_performance += [None] * (max_length - len(period_of_performance))
    total_project_price += [None] * (max_length - len(total_project_price))
    deliverables += [None] * (max_length - len(deliverables))

    # Create DataFrame
    df = pd.DataFrame({
        "Project Scope": project_scope,
        "Period of Performance": period_of_performance,
        "Total Project Price": total_project_price,
        "Deliverables": deliverables
    })

    # Save DataFrame to Excel
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Extracted Data')

    excel_buffer.seek(0)
    return excel_buffer

st.title('Document Intelligence Tool')
uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

if uploaded_file is not None:
    st.write('PDF file uploaded successfully.')
    st.write('Processing...')
    excel_buffer = process_pdf(uploaded_file)
    
    if excel_buffer:
        st.success('Excel file generated successfully!')
        st.download_button(
            label="Download Excel file",
            data=excel_buffer,
            file_name="extracted_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
