import streamlit as st
import openai
import json
from pptx import Presentation
from pptx.util import Inches
from dotenv import load_dotenv
import os
from st_files_connection import FilesConnection
import pandas as pd
import re
import ast

# Load environment variables from a .env file if present
load_dotenv()
#path = os.path.dirname(__file__)
# Set your OpenAI API key
openai.api_key = st.text_input("OpenAI API Key", type="password")

# Title of the Streamlit app
st.title('Slide Content Generator')

# Function to generate slide content
def generate_slide_content(topic, content):

    conn = st.connection('s3', type=FilesConnection)
    st.write("conn obtained")
    
    df = conn.read("fbc-hackathon-test/growth.csv", input_format="csv", ttl=600)
    #st.write("df obtained")
    median = conn.read("fbc-hackathon-test/Network_Median.csv", input_format="csv", ttl=600)
    #st.table(df)
    # Print results.
    #st.write(median.to_dict())
    #for row in median.itertuples():
    #    st.write(f"{row}")
    
    prompt_txt = f"Wait for user input to return a response. Use this data to generate the output:\n\n{df.to_string()}"
    prompt = f"You are a helpful assistant that generates an executive summary of Franchise's performance metrics. For each comma separated Franchise number in the list {topic} return output as a Python dictionary with the following keys: First Name & Last Name as Franchisee, NetworkPerformancePartner as FBC, State as DO, Weighted Score, Rank, Current Billable hours, Previous year billable hours, Growth hours %, Current total revenue, Previous year total revenue. Then calculate aggregate metrics for all Franchises and return out as a python dictionary. Lastly summarizekey insights on Franchise metrics. Do not return anything else."

    # Use ChatCompletion with the new model and API method
    response = openai.chat.completions.create(
        model="gpt-3.5-turbo",  # Specify the model
        messages=[
            {"role": "system", "content": prompt_txt},
            {"role": "user", "content": prompt}
        ],
        temperature=0.7,
    )
    generated_text = response.choices[0].message.content
    match = re.search(r'\\{.*?\\}', generated_text, re.DOTALL)
    dictionary = None
    if match:
        try:
            # Try to convert substring to dict
            dictionary = ast.literal_eval(match.group())
        except (ValueError, SyntaxError):
            # Not a dictionary
            return None
    st.write(dictionary)
    return generated_text

# Function to create a PowerPoint presentation
def create_presentation(topic, slide_content):
    #pptx = path + '//' + 'template.pptx'
    prs = Presentation()
    title_slide_layout = prs.slide_layouts[0]
    bullet_slide_layout = prs.slide_layouts[1]

    # Title slide
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = topic
    subtitle.text = "Generated using OpenAI and Streamlit"

    for key, value in data.items():
        if isinstance(value, list):
            print(f"{key}:")
            for item in value:
                slide = prs.slides.add_slide(bullet_slide_layout)
                shapes = slide.shapes
                title_shape = shapes.title
                body_shape = shapes.placeholders[1]
                title_shape.text = f"{key}"
                tf = body_shape.text_frame
                for sub_key, sub_value in item.items():
                    print(f"  {sub_key}: {sub_value}")
                    p = tf.add_paragraph()
                    p.text+= f"  {sub_key}: {sub_value} \n\n"
                #print()  # Line break between items
        elif isinstance(value, dict):
            print(f"{key}:")
            slide = prs.slides.add_slide(bullet_slide_layout)
            shapes = slide.shapes
            title_shape = shapes.title
            body_shape = shapes.placeholders[1]
            title_shape.text = f"{key}"
            tf = body_shape.text_frame
            for sub_key, sub_value in value.items():
                print(f"  {sub_key}: {sub_value}")
                p = tf.add_paragraph()
                p.text+= f"  {sub_key}: {sub_value} \n\n"
        else:
            print(f"{key}: {value}")
            slide = prs.slides.add_slide(bullet_slide_layout)
            shapes = slide.shapes
            title_shape = shapes.title
            body_shape = shapes.placeholders[1]
            title_shape.text = f"{key}"
            tf = body_shape.text_frame
            p = tf.add_paragraph()
            p.text = f"  {key}: {value} \n\n"
    
    # Save the presentation
    file_path = "generated_presentation.pptx"
    prs.save(file_path)
    return file_path

# Streamlit input fields
topic = st.text_input("Enter the Franchise number:")
content = st.text_area("Enter the themes for the slides:")

# Generate button
if st.button("Generate Slide Content"):
    if topic and content:
        generated_content = generate_slide_content(topic, content)
        st.subheader("Generated Slide Content:")
        st.write(generated_content)
        
        # Create and offer download of the PowerPoint presentation
        file_path = create_presentation(topic, generated_content)
        with open(file_path, "rb") as file:
            btn = st.download_button(
                label="Download PowerPoint Presentation",
                data=file,
                file_name=file_path,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
    else:
        st.error("Please enter both the topic and content.") 
