import streamlit as st
import openai
import json
from pptx import Presentation
from pptx.util import Inches
from dotenv import load_dotenv
import os
from st_files_connection import FilesConnection
import pandas as pd

# Load environment variables from a .env file if present
load_dotenv()

# Set your OpenAI API key
openai.api_key = st.text_input("OpenAI API Key", type="password")

# Title of the Streamlit app
st.title('Slide Content Generator')

# Function to generate slide content
def generate_slide_content(topic, content):

    conn = st.connection('s3', type=FilesConnection)
    st.write("conn obtained")
    df = conn.read("fbc-hackathon-test/Operations ScoreCard - UX.csv", input_format="csv", ttl=600)
    #st.write("df obtained")
    #st.table(df)
    # Print results.
    #for row in df.itertuples():
    #st.write(f"{row}")
    
    prompt = f"Generate slide ideas for {topic}:\n\n{df.to_string()}"
    
    # Use ChatCompletion with the new model and API method
    response = openai.chat.completions.create(
        model="gpt-3.5-turbo",  # Specify the model
        messages=[
            {"role": "system", "content": "You are a helpful assistant. Use the data from CSV to summarize Billable, RPN & Growth metrics for a Franchise."},
            {"role": "user", "content": prompt}
        ],
        temperature=0.7,
    )
    generated_text = response.choices[0].message.content
    return generated_text

# Function to create a PowerPoint presentation
def create_presentation(topic, slide_content):
    prs = Presentation()
    title_slide_layout = prs.slide_layouts[0]
    bullet_slide_layout = prs.slide_layouts[1]

    # Title slide
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = topic
    subtitle.text = "Generated using OpenAI and Streamlit"

    # Parse the slide content
    slides = slide_content.split('\n\n')
    for slide in slides:
        lines = slide.split('\n')
        slide_title = lines[0].replace('Title: ', '')
        slide_content = '\n'.join(lines[1:]).replace('- ', '')

        # Title slide
        #slide = prs.slides.add_slide(title_slide_layout)
        #title = slide.shapes.title
        #title.text = slide_title

        # Content slide
        slide = prs.slides.add_slide(bullet_slide_layout)
        shapes = slide.shapes
        title_shape = shapes.title
        body_shape = shapes.placeholders[1]
        
        title_shape.text = shapes.title
        tf = body_shape.text_frame
        for content_line in slide_content.split('\n'):
            p = tf.add_paragraph()
            p.text = content_line

    # Save the presentation
    file_path = "generated_presentation.pptx"
    prs.save(file_path)
    return file_path

# Streamlit input fields
topic = st.text_input("Enter the topic of the slides:")
content = st.text_area("Enter the content or bullet points for the slides:")

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


