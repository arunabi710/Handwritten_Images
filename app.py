import streamlit as st
import io
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
from openpyxl import Workbook

# Azure Form Recognizer credentials
endpoint = "https://arun-document.cognitiveservices.azure.com/"
credential = AzureKeyCredential("613a3ef0c05a465b96b6614fa026e162")
document_analysis_client = DocumentAnalysisClient(endpoint, credential)
model_id = "Sow-Template-Processing"

def process_pdf(file):
    # Process the uploaded PDF
    document = file.read()
    
    poller = document_analysis_client.begin_analyze_document(model_id, document)
    result = poller.result()

    # Create Excel workbook and write data
    workbook = Workbook()
    sheet = workbook.active
    
    # Define headers and row data in the desired order
    headers = ["Project Scope", "Period of Performance", "Total Project Price", "Deliverables"]
    sheet.append(headers)
    row_data = [
        result.documents[0].fields.get("Project Scope", {}).value,
        result.documents[0].fields.get("Period of Performance", {}).value,
        result.documents[0].fields.get("Total Project Price", {}).value,
        result.documents[0].fields.get("Deliverables", {}).value
    ]
    sheet.append(row_data)

    # Save the workbook to a bytes buffer
    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    
    return buffer

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