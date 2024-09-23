import streamlit as st
import pandas as pd
from io import BytesIO
import json
from azure.core.credentials import AzureKeyCredential
from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.ai.formrecognizer import DocumentAnalysisClient
from PIL import Image


# config_path = 'config.json'
# with open(config_path, 'r') as config_file:
#     config = json.load(config_file)

azure_document_api_key = st.secrets['azure_document_api_key']
azure_document_endpoint = st.secrets['azure_document_endpoint']
custom_model_id = st.secrets['custom_model_id']


document_analysis_client = DocumentAnalysisClient(
    endpoint=azure_document_endpoint,  
    credential=AzureKeyCredential(azure_document_api_key)
)

document_intelligence_client = DocumentIntelligenceClient(
    endpoint=azure_document_endpoint, 
    credential=AzureKeyCredential(azure_document_api_key)
)


def analyze_invoice(uploaded_file):
    try:
        file_stream = BytesIO(uploaded_file.getvalue())
        poller = document_analysis_client.begin_analyze_document("prebuilt-invoice", file_stream)
        result = poller.result()
        return result
    except Exception as e:
        st.error(f"Error processing invoice: {str(e)}")
        return None
    

def analyze_custom_model(uploaded_file):
    try:
        file_stream = BytesIO(uploaded_file.getvalue())
        if uploaded_file.name.lower().endswith('.pdf'):
            content_type = "application/pdf"
        elif uploaded_file.name.lower().endswith(('.jpg', '.jpeg')):
            content_type = "image/jpeg"
        elif uploaded_file.name.lower().endswith('.png'):
            content_type = "image/png"
        else:
            st.error("Unsupported file type. Please upload a PNG, JPG, JPEG, or PDF file.")
            return None

        poller = document_intelligence_client.begin_analyze_document(
            custom_model_id, 
            file_stream, 
            content_type=content_type
        )
        result = poller.result()
        return result
    except Exception as e:
        st.error(f"Error processing invoice layout: {str(e)}")
        return None
    

def layout_invoice(uploaded_file):
    try:
        file_stream = BytesIO(uploaded_file.getvalue())
        if uploaded_file.name.lower().endswith('.pdf'):
            content_type = "application/pdf"
        elif uploaded_file.name.lower().endswith(('.jpg', '.jpeg')):
            content_type = "image/jpeg"
        elif uploaded_file.name.lower().endswith('.png'):
            content_type = "image/png"
        else:
            st.error("Unsupported file type. Please upload a PNG, JPG, JPEG, or PDF file.")
            return None

        poller = document_intelligence_client.begin_analyze_document(
            "prebuilt-layout", 
            file_stream, 
            content_type=content_type
        )
        result = poller.result()
        return result
    except Exception as e:
        st.error(f"Error processing invoice layout: {str(e)}")
        return None

def extract_field_data(invoice_data, custom_data=None):
    all_field_data = []

    for doc in invoice_data.documents:
        for field_name, field in doc.fields.items():
            # Extract value and confidence from each field
            value = field.content if hasattr(field, 'content') and field.content else 'N/A'
            confidence = field.confidence if hasattr(field, 'confidence') else 'N/A'
            all_field_data.append({'Key': field_name, 'Value': value, 'Confidence': confidence})

    if custom_data:
        for doc in custom_data.documents:
            for field_name, field in doc.fields.items():
                value = field.content if hasattr(field, 'content') and field.content else 'N/A'
                confidence = field.confidence if hasattr(field, 'confidence') else 'N/A'
                if value != 'N/A':
                    # Add custom fields only if they don't already exist
                    if field_name not in [data['Key'] for data in all_field_data]:
                        all_field_data.append({'Key': field_name, 'Value': value, 'Confidence': confidence})
            
    
    return pd.DataFrame(all_field_data)



def extract_table_data(layout_data):
    tables_list = []
    if hasattr(layout_data, 'tables'):
        for table in layout_data.tables:
            table_data = []
            for cell in table.cells:
                if len(table_data) <= cell.row_index:
                    table_data.extend([{} for _ in range(cell.row_index + 1 - len(table_data))])
                column_header = f"Column {cell.column_index}" 
                table_data[cell.row_index][column_header] = cell.content
            if table_data:
                df = pd.DataFrame(table_data)
                tables_list.append(df)
    return tables_list


def create_excel(fields_df, tables_list):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Add the invoice fields to the first sheet
        fields_df.to_excel(writer, index=False, sheet_name="Invoice Fields")

        # Add each table to a separate sheet
        for idx, table_df in enumerate(tables_list):
            sheet_name = f"Table_{idx + 1}"
            table_df.to_excel(writer, index=False, sheet_name=sheet_name)
        
    output.seek(0)
    return output

# Streamlit Application

st.set_page_config(
    page_title = "MyPay Document Scanner",
    page_icon = "page_icon.jpeg",
    layout="wide",
    initial_sidebar_state="expanded"
)

img=Image.open("logo.png")
st.image(img, width=100)

st.title("MyPay Document Scanner")

uploaded_file = st.file_uploader("Upload your invoice (PDF, JPG, PNG)", type=["pdf", "jpg", "jpeg", "png"])

if uploaded_file:
    st.write("Extracting data from the invoice...")
    
    invoice_data = analyze_invoice(uploaded_file)
    custom_data = analyze_custom_model(uploaded_file)
    layout_data = layout_invoice(uploaded_file)

    # Extract fields and tables
    if invoice_data and invoice_data.documents:
        fields_df = extract_field_data(invoice_data, custom_data)
        if not fields_df.empty:
            st.write("Extracted Field Data:")
            st.data_editor(fields_df, num_rows = "dynamic")
    
    if layout_data:
        tables_list = extract_table_data(layout_data)
        if tables_list:
            st.write("Extracted Tables:")
            for idx, table_df in enumerate(tables_list):
                st.write(f"Table {idx + 1}:")
                tables_list[idx] = st.data_editor(table_df, num_rows="dynamic")
                # st.dataframe(table_df)

    if st.button('Finalize the Edits'):
        if invoice_data and layout_data:
            excel_file = create_excel(fields_df, tables_list)
            st.download_button(
                label="Download Excel file",
                data=excel_file,
                file_name="extracted_invoice_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("Please upload an invoice to extract data.")
