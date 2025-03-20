import streamlit as st
import pandas as pd
import openai
from dotenv import load_dotenv
import os

# Load environment variables from .env file
load_dotenv()

# Set up OpenAI API key
client = openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

def main():
    st.title("Excel Q&A with OpenAI")
    
    # File uploader
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])
    
    if uploaded_file is not None:
        # Read the Excel file
        df = pd.read_excel(uploaded_file)
        st.write("Data Preview:")
        # Convert dataframe to a compatible format
        df = df.astype(str)
        st.dataframe(df)
        
        # User question input
        question = st.text_input("Ask a question about the data:")
        
        if question:
            # Call OpenAI API using the new format
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant."},
                    {"role": "user", "content": f"Data: {df.to_string(index=False)}\n\nQuestion: {question}"}
                ],
                max_tokens=150
            )
            answer = response.choices[0].message.content.strip()
            st.write("Answer:", answer)

if __name__ == "__main__":
    main()
