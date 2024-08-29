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
import io
import csv

# Load environment variables from a .env file if present
load_dotenv()
path = os.path.dirname(__file__)
# Set your OpenAI API key
openai.api_key = st.text_input("OpenAI API Key", type="password")

# Title of the Streamlit app
st.title('Slide Content Generator')

# Function to generate slide content
def generate_slide_content(topic, content):

    conn = st.connection('s3', type=FilesConnection)
    #st.write("conn obtained")
    
    df = conn.read("fbc-hackathon-test/growth.csv", input_format="csv", ttl=600)
    csv_file = io.StringIO(df.to_string())

    # Read the CSV data into a list of dictionaries
    reader = csv.DictReader(csv_file)
    dict_list = [row for row in reader]

    # Now, dict_list contains the CSV content as a list of dictionaries
    #for item in dict_list:
    first_key_value_pair = next(iter(dict_list[0].items()))

    # Now, first_key_value_pair contains the first key-value pair of the first row
    st.write(first_key_value_pair)
        
    #st.write("df obtained")
    median = conn.read("fbc-hackathon-test/Network_Median.csv", input_format="csv", ttl=600)
    #st.table(df)
    # Print results.
    #st.write(median.to_dict())
    #for row in median.itertuples():
    #    st.write(f"{row}")
    st.write(f"Raw CSV: {df.to_string()}")
    #st.write(f"Raw dict: {df.to_dict()}")
    prompt_txt = f"Wait for user input to return a response. Use this data to generate the output as a single python dictionary:\n\n{df.to_string()}"
    prompt = f"You are a helpful assistant that generates an executive summary of Franchise's performance metrics. For each comma separated Franchise number in the list {topic} return all the data as a list of Python dict object. Then calculate aggregate metrics for all Franchises and return output as a python dict. Lastly summarizekey insights on Franchise metrics. Return all output as a single python dict object. Do not return anything else."

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
    st.write(f"Response: {generated_text}")
    return generated_text

# function to replace text in pptx first slide with selected filters
def replace_text(replacements, shapes):
    """function to replace text on a PowerPoint slide. Takes dict of {match: replacement, ... } and replaces all matches"""
    for shape in shapes:
        for match, replacement in replacements.items():
            if shape.has_text_frame:
                if (shape.text.find(match)) != -1:
                    text_frame = shape.text_frame
                    for paragraph in text_frame.paragraphs:
                        whole_text = "".join(run.text for run in paragraph.runs)
                        whole_text = whole_text.replace(str(match), str(replacement))
                        for idx, run in enumerate(paragraph.runs):
                            if idx != 0:
                                p = paragraph._p
                                p.remove(run._r)
                        if bool(paragraph.runs):
                            paragraph.runs[0].text = whole_text

# Function to create a PowerPoint presentation
def create_presentation(topic, slide_content):
    pptx = path + '//' + 'template.pptx'
    prs = Presentation(pptx)
    #title_slide_layout = prs.slide_layouts[0]
    bullet_slide_layout = prs.slide_layouts.get_by_name('Purple_Circle_Corners')

    # Title slide
    #slide = prs.slides.add_slide(title_slide_layout)
    #title = slide.shapes.title
    #subtitle = slide.placeholders[1]
    #title.text = topic
    #subtitle.text = "Generated using OpenAI and Streamlit"
    
    slide_content = ast.literal_eval(slide_content)
    st.write(isinstance(slide_content, dict))

    details_dict = {
    'Franchisee': 'Franchisee',
    'NetworkPerformancePartner': 'FBC',
    'Region': 'DO',
    'WeightedScore': 'Your Total Score',
    'aggregate_metrics': 'Aggregate Metrics',    
    'key_insights': 'Key Insights'
    }

    owner = []
    for key, value in slide_content.items():
        if isinstance(value, list):
            #print(f"{key}:")
            for item in value:
                slide = prs.slides.add_slide(bullet_slide_layout)
                shapes = slide.shapes
                title_shape = shapes.title
                st.write(item.items())
                title_shape.text = f"Franchise {item[.items()['Number']} - {item.items()['FirstName']} {item.items()['LastName']}"
                owner.append(f"{item['FirstName']} {item['LastName']}")
                body_shape = shapes.placeholders[1]
                tf = body_shape.text_frame
                for sub_key, sub_value in item.items():
                    if sub_key in details_dict:
                        p = tf.add_paragraph()
                        p.text+= f"  {details_dict[sub_key]}: {sub_value}\n"
                #print()  # Line break between items
        elif isinstance(value, dict):
            st.write(f" when does it come here - {key}")
            slide = prs.slides.add_slide(bullet_slide_layout)
            shapes = slide.shapes
            title_shape = shapes.title
            body_shape = shapes.placeholders[1]
            tf = body_shape.text_frame
            title_shape.text = f"{details_dict[key]}"
            for sub_key, sub_value in value.items():
                if sub_key in details_dict:
                    p = tf.add_paragraph()
                    p.text+= f"  {details_dict[sub_key]}: {sub_value}\n"
        else:
            #print(f"{key}: {value}")
            slide = prs.slides.add_slide(bullet_slide_layout)
            shapes = slide.shapes
            title_shape = shapes.title
            title_shape.text = f"{key} - {value}"
            body_shape = shapes.placeholders[1]
            tf = body_shape.text_frame
            if key in details_dict:
                p = tf.add_paragraph()
                p.text+= f"  {details_dict[key]}: {value}\n"
    owner = list(set(owner))
    #st.write(owner)
    # Convert the array to a comma-separated string
    comma_separated_string = ", ".join(owner)
    #st.write(f"comma separated unique list: {comma_separated_string}")
    first_slide = prs.slides[0]
    shapes_1 = []

    # create lists with shape objects
    for shape in first_slide.shapes:
        shapes_1.append(shape)

    # initiate a dictionary of placeholders and values to replace
    replaces_1 = {
        '{o}': comma_separated_string}
    replace_text(replaces_1, shapes_1)
    
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
        #st.write(generated_content)
        
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
