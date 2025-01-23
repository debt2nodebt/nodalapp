import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

# Function to fetch details from the Excel file
def fetch_emails_from_excel(df, bank_names):
    result = []
    for bank in bank_names:
        row = df[df['Bank Name'].str.lower() == bank.lower()]
        if not row.empty:
            details = {
                'Bank Name': bank,
                'Customer Email': row['Customer Email'].values[0] if not pd.isna(row['Customer Email'].values[0]) else '',
                'Nodal Email': row['Nodal Email'].values[0] if not pd.isna(row['Nodal Email'].values[0]) else '',
                'Grievance Email': row['Grievance Email'].values[0] if not pd.isna(row['Grievance Email'].values[0]) else ''
            }
        else:
            details = {'Bank Name': bank, 'Customer Email': '', 'Nodal Email': '', 'Grievance Email': ''}
        result.append(details)
    return result

# Function to create a Word document
def create_word_file(bank_details):
    doc = Document()
    doc.add_heading("Banks Email (Nodal)", level=1)

    for details in bank_details:
        doc.add_paragraph(f"Bank Name: {details['Bank Name']}")
        doc.add_paragraph(f"Customer Email: {details['Customer Email']}")
        doc.add_paragraph(f"Nodal Email: {details['Nodal Email']}")
        doc.add_paragraph(f"Grievance Email: {details['Grievance Email']}")
        doc.add_paragraph("")  # Add a blank line after each bank's details

    # Save the Word file to a BytesIO object
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# Streamlit App
def main():
    st.title("Nodal Email Generator")

    # File uploader
    uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx'])
    
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        st.success("File loaded successfully!")
        st.dataframe(df.head())  # Show preview of the uploaded file

        # Input for bank names
        bank_names_input = st.text_area("Enter Bank Names (comma-separated)", "")
        if bank_names_input:
            bank_names = [name.strip() for name in bank_names_input.split(',')]

            # Generate the document
            if st.button("Generate Word File"):
                bank_details = fetch_emails_from_excel(df, bank_names)
                if bank_details:
                    word_file = create_word_file(bank_details)
                    st.download_button(
                        label="Download Word File",
                        data=word_file,
                        file_name="Banks_Email_Nodal.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
    else:
        st.warning("Please upload the Excel file to proceed.")

if __name__ == "__main__":
    main()
