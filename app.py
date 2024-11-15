import os
import docx
import streamlit as st
import openai
import pandas as pd
from io import BytesIO
from dotenv import load_dotenv
import re

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
    system_prompt = f"""You are a highly skilled Business Analyst/System Analyst (BA/SA) assistant, tasked with generating questions for system integration testing (SIT) based on the document content. The document may contain regular text and unrolled table data. Always use Vietnamese. 

### Instructions:
1. **Generate Questions**: For each section, heading, or specific term in the document, create questions based solely on the content provided, with the exact wording and phrases from the document. Do not paraphrase or modify any text.
2. **Document Name and Context**: Include the document name in the question if available, along with any descriptors such as "22 pages, 25 pages, 30 pages, etc." for added context.
3. **Unrolled Table Data**: For unrolled tables, combine headings, subheadings, and content to form a coherent question that maintains the intended context and meaning.
4. **Special Sections**: Identify any terms or sections enclosed in brackets, such as [STT], [Hoạt động], [VAS], [BC TCQT]. For each detected [Hoạt động], use the question format:
   "Hoạt động '[Hoạt động]' thuộc VAS nào và BC TCQT nào?"
5. **Reference Paragraph**: After each question, include the exact paragraph or relevant content from the document as a reference, labeled 'Reference Paragraph:'. Ensure this reference preserves the original wording and context.

### Format:
For each item, your output should be structured as follows:
- **Question**: Write a specific question based on the content and context of the document.
- **Reference Paragraph**: Provide the exact paragraph or relevant text that the question references, labeled 'Reference Paragraph:'.

### Examples:
- **Question**: "Hoạt động 'Thu từ nghiệp vụ bảo lãnh' thuộc VAS nào và BC TCQT nào?"
  **Reference Paragraph**: "Hoạt động này bao gồm tất cả các khoản thu liên quan đến nghiệp vụ bảo lãnh mà VPBank thực hiện cho khách hàng."

- **Question**: "Yêu cầu cho phần tiêu đề 'Thông tin khách hàng' dưới heading 'Đăng ký người dùng' là gì?"
  **Reference Paragraph**: "Phần 'Thông tin khách hàng' dưới mục 'Đăng ký người dùng' yêu cầu thu thập các thông tin cá nhân bao gồm tên, địa chỉ, và thông tin liên hệ của khách hàng."

- **Question**: "[STT] 5: Hoạt động 'Thu nhập từ góp vốn mua cổ phần' thuộc VAS nào và BC TCQT nào?"
  **Reference Paragraph**: "Đây là các khoản thu nhập phát sinh từ các hoạt động góp vốn, đầu tư vào cổ phần của doanh nghiệp khác mà VPBank tham gia."

### Additional Notes:
- Keep each question directly relevant to the document content, ensuring all extracted questions and reference paragraphs are specific and retain the document's original terminology.
- When handling unrolled tables, ensure that combined elements form a natural and precise question, followed by the exact unrolled content in the Reference Paragraph section.
"""

    
    user_prompt = f"""Document name: {doc_name}\n\nDocument content:\n\n{doc_content}\n\nGenerate SIT questions from the above content."""

    response = openai.ChatCompletion.create(
        model=openai_model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ]
    )
    return response.choices[0].message['content']

# Function to convert questions into a pandas DataFrame
def create_sit_question_dataframe(response_text):
    # Regular expression to match questions and reference paragraphs without the labels
    pattern = r"- \*\*Question\*\*: \"(.*?)\"\s*\*\*Reference Paragraph\*\*: \"(.*?)\""
    matches = re.findall(pattern, response_text, re.DOTALL)
    
    # Lists to store extracted questions and reference paragraphs
    question_list = []
    reference_list = []
    
    for match in matches:
        question = match[0].strip()
        reference_paragraph = match[1].strip()
        question_list.append(question)
        reference_list.append(reference_paragraph)
    
    # Create DataFrame without the labels
    data = {'Question': question_list, 'Reference Paragraph': reference_list}
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
        st.stop()
    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Please upload a DOCX file to generate SIT questions.")
