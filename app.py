import os
import docx
import streamlit as st
import openai
import pandas as pd
from io import BytesIO
from dotenv import load_dotenv

# Load environment variables in a file called .env
load_dotenv()
openai.api_key = os.getenv('OPENAI_API_KEY')
openai_model = os.getenv('MODEL_NAME')  # Using your specified model name

# Function to extract text from a DOCX file
def extract_text_from_docx(docx_file):
    doc = docx.Document(docx_file)
    full_text = []
    for paragraph in doc.paragraphs:
        full_text.append(paragraph.text)
    return '\n'.join(full_text)

# Function to generate SIT questions based on DOCX content, including document name
def generate_sit_questions(doc_content, doc_name):
    system_prompt = """You are an assistant tasked with generating questions for \
system integration testing (SIT) based solely on the contents of the provided \
document. You always use Vietnamese. You must not paraphrase or modify any text. Retain original wording and \
phrases exactly as they appear in the document. For example, phrases like "thẻ visa \
lady mastercard" or "Nghị định 08/2018" must remain unchanged.
    
Generate questions based on this content for SIT purposes. Remember to extract the document name, if there are any additional words like "22 trang, 25 trang, 30 trang,...." and append to the question to keep the context. For example: 
""
"""
    
    user_prompt = f"Document name: {doc_name}\n\nDocument content:\n\n{doc_content}\n\nGenerate SIT questions from the above content."

    response = openai.ChatCompletion.create(
        model=openai_model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ]
    )
    return response.choices[0].message['content']

# Function to convert questions into a pandas DataFrame
def create_sit_question_dataframe(questions):
    question_list = questions.strip().split('\n')  # Assuming each question is separated by a new line
    data = {'STT': range(1, len(question_list) + 1), 'Câu hỏi': question_list}
    df = pd.DataFrame(data)
    return df

# Function to convert DataFrame to Excel and return as downloadable file
def convert_df_to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='SIT Questions')
    writer.close()  # Ensure to close the writer
    processed_data = output.getvalue()
    return processed_data

# Streamlit app interface
st.title("SIT Question Generator from DOCX")

# Upload DOCX file
docx_file = st.file_uploader("Upload a DOCX file", type="docx")

# If a DOCX file is uploaded, read and process it
if docx_file is not None:
    try:
        # Extract the text from the DOCX file
        doc_content = extract_text_from_docx(docx_file)
        doc_name = docx_file.name  # Get the name of the DOCX file
        st.write(f"Extracting content from the DOCX file: {doc_name}")
        
        # Generate SIT questions
        sit_questions = generate_sit_questions(doc_content, doc_name)
        st.write("Generated SIT Questions:")
        
        # Create DataFrame from the questions
        df = create_sit_question_dataframe(sit_questions)
        st.dataframe(df)  # Display the DataFrame
        
        # Convert DataFrame to Excel and provide download link
        excel_data = convert_df_to_excel(df)
        st.download_button(label="Download as Excel", 
                           data=excel_data, 
                           file_name="sit_questions.xlsx", 
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Please upload a DOCX file to generate SIT questions.")
