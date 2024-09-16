import streamlit as st
import pandas as pd
from io import BytesIO
from azure.core.credentials import AzureKeyCredential
from azure.ai.documentintelligence import DocumentIntelligenceClient

# Load configuration
config_path = 'config.json'
with open(config_path, 'r') as config_file:
    config = json.load(config_file)

azure_document_api_key = config['azure_document_api_key']
azure_document_endpoint = config['azure_document_endpoint']

# Initialize Azure Document Intelligence Client
document_intelligence_client = DocumentIntelligenceClient(
    endpoint=azure_document_endpoint,
    credential=AzureKeyCredential(azure_document_api_key)
)

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

        # Call the prebuilt-layout model correctly without the 'document' keyword
        poller = document_intelligence_client.begin_analyze_document(
            model_id="prebuilt-layout",  # Use Layout model
            document=file_stream.read(),  # Pass the binary stream directly
            content_type=content_type
        )
        
        result = poller.result()
        return result
    except Exception as e:
        st.error(f"Error processing invoice layout: {str(e)}")
        return None

def extract_table_data(document):
    tables_list = []
    
    if hasattr(document, 'tables'):
        for table_idx, table in enumerate(document.tables):
            table_data = []
            for cell in table.cells:
                while len(table_data) <= cell.row_index:
                    table_data.append({})
                # Build the table as rows and columns
                column_header = f"Column {cell.column_index + 1}"
                table_data[cell.row_index][column_header] = cell.content
            df = pd.DataFrame(table_data)
            tables_list.append(df)
    
    if tables_list:
        return pd.concat(tables_list, ignore_index=True)  # Combine all extracted tables into a single DataFrame
    else:
        return pd.DataFrame()

# Streamlit UI Setup
st.title("Extract Tables from Invoice")

uploaded_file = st.file_uploader("Upload your invoice (PDF, JPG, JPEG, PNG)", type=["pdf", "jpg", "jpeg", "png"])

if uploaded_file:
    st.write("Extracting table data from the invoice...")
    
    layout_data = layout_invoice(uploaded_file)

    if layout_data and hasattr(layout_data, 'tables'):
        table_df = extract_table_data(layout_data)
        
        if not table_df.empty:
            st.write("Extracted Table Data:")
            st.dataframe(table_df)

            # Option to download the data as Excel
            def create_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                output.seek(0)
                return output

            if st.button('Download Extracted Table Data as Excel'):
                excel_file = create_excel(table_df)
                st.download_button(
                    label="Download Excel File",
                    data=excel_file,
                    file_name="extracted_table_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.write("No tables found in the document.")
    else:
        st.write("No data could be extracted. Please try again.")
else:
    st.info("Please upload an invoice to extract table data.")
