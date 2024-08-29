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
def get_franchise_data(topic):
    conn = st.connection('s3', type=FilesConnection)
    #st.write("conn obtained")
    
    df = conn.read("fbc-hackathon-test/growth.csv", input_format="csv", ttl=600) 
    df = df.transpose()
    df.columns = df.iloc[0]  # Use the first row as the header
    df = df.drop(df.index[0])  # Drop the first row since it is now the header
    df = df.to_dict()
    #st.write(df)
    keys_to_keep = topic.split(',')
    keys_to_keep = [key.strip() for key in topic.split(',')]
    #st.write(keys_to_keep)
    filtered_dict = {}
    for data in df:
        if str(data) in keys_to_keep:
            #st.write("should come here multiple times")
            #st.write(data)
            #st.write(df.get(data))
            #st.write(df[data])
            filtered_dict[str(data)] = df[data]
    return filtered_dict
    
# Function to generate slide content
def generate_slide_content(content):
    prompt_txt = f"Wait for user input to return a response. Use this data to generate the output:\n\n{str(content)}"
    prompt = f"You are a helpful assistant that generates an executive summary of Franchise's performance metrics. Calculate aggregate metrics for given Franchises and return output as a python dict with key as aggregate_metrics. Then summarize 3 Key insights comparing metrics between previous and current year as a python dict object with key as key_insights. Return output as a valid python dict object. Do not return anything else."

    # Use ChatCompletion with the new model and API method
    response = openai.chat.completions.create(
        model="gpt-4o-2024-08-06",  # Specify the model
        messages=[
            {"role": "system", "content": prompt_txt},
            {"role": "user", "content": prompt}
        ],
        temperature=0.7,
    )
    generated_text = response.choices[0].message.content
    generated_text = generated_text.removeprefix("```python")
    st.write(isinstanceof(generated_text, dict))
    st.write(f"Response: {generated_text}")
    return ast.literal_eval(generated_text)

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
def create_presentation(franchise_data, slide_content):
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
    
    #slide_content = ast.literal_eval(slide_content)
    #st.write(isinstance(slide_content, dict))
    #st.write(slide_content)
    details_dict = {
    'Franchisee': 'Franchisee',
    'NetworkPerformancePartner': 'FBC',
    'Region': 'DO',
    'WeightedScore': 'Your Total Score',
    'aggregate_metrics': 'Aggregate Metrics',    
    'key_insights': 'Key Insights',
    'AggregateMetrics': 'Aggregate Metrics',    
    'KeyInsights': 'Key Insights'
    }

    owner = []
    aggregate_metrics = {}
    key_insights = {}
    #st.write(f"so far so good")
    #st.write(isinstance(slide_content, dict))
    for item in slide_content:
        st.write(item)
    if 'aggregate_metrics' in slide_content:
        aggregate_metrics = slide_content['aggregate_metrics']
    if 'AggregateMetrics' in slide_content:
        aggregate_metrics = slide_content['AggregateMetrics']
    if 'key_insights' in slide_content:
        key_insights = slide_content['key_insights']
    if 'KeyInsights' in slide_content:
        key_insights = slide_content['KeyInsights']
    #st.write(franchise_data)
    #st.write(key_insights)
    st.write(aggregate_metrics)

    for franchise in franchise_data:
        slide = prs.slides.add_slide(bullet_slide_layout)
        shapes = slide.shapes
        title_shape = shapes.title
        body_shape = shapes.placeholders[1]
        tf = body_shape.text_frame
        ind_fran = franchise_data[str(franchise)]
        st.write(franchise)
        st.write(ind_fran['FirstName'])
        owner_name = f"Franchise {franchise} - {ind_fran['FirstName']} {ind_fran['LastName']}"
        st.write(owner_name)
        owner.append(owner_name)
        title_shape.text = owner_name
        for k in ind_fran:
            if k in details_dict:
                p = tf.add_paragraph()
                p.text+= f"  {details_dict[k]}: {ind_fran[k]}\n\n"
            
    owner = list(set(owner))
    st.write(owner)
    # Convert the array to a comma-separated string
    comma_separated_string = ", ".join(owner)
    st.write(f"comma separated unique list: {comma_separated_string}")
    first_slide = prs.slides[0]
    shapes_1 = []

    #Aggregate Metrics
    slide = prs.slides.add_slide(bullet_slide_layout)
    shapes = slide.shapes
    title_shape = shapes.title
    title_shape.text = f"The Enterprise Journey of {owner}"
    body_shape = shapes.placeholders[1]
    tf = body_shape.text_frame
    for k in aggregate_metrics:
        p = tf.add_paragraph()
        p.text+= f"  {k}: {aggregate_metrics[k]}\n\n"

    #Key Insights
    slide = prs.slides.add_slide(bullet_slide_layout)
    shapes = slide.shapes
    title_shape = shapes.title
    title_shape.text = f"Key Insights"
    body_shape = shapes.placeholders[1]
    tf = body_shape.text_frame
    for k in key_insights:
        p = tf.add_paragraph()
        p.text+= f"  {k}: {key_insights[k]}\n\n\n"
        
    # create lists with shape objects
    for shape in first_slide.shapes:
        shapes_1.append(shape)

    # initiate a dictionary of placeholders and values to replace
    replaces_1 = {
        '{o}': comma_separated_string}
    replace_text(replaces_1, shapes_1)

    st.write(key_insights)
    
    # Save the presentation
    file_path = "generated_presentation.pptx"
    prs.save(file_path)
    return file_path

# Streamlit input fields
topic = st.text_input("Enter the Franchise number:")
content = st.text_area("Enter the themes for the slides:")

# Generate button
if st.button("Generate Slide Content"):
    if topic:
        franchise_data = get_franchise_data(topic)
        #st.write(franchise_data)
        generated_content = generate_slide_content(franchise_data)
        st.subheader("Generated Slide Content:")
        
        # Create and offer download of the PowerPoint presentation
        file_path = create_presentation(franchise_data, generated_content)
        with open(file_path, "rb") as file:
            btn = st.download_button(
                label="Download PowerPoint Presentation",
                data=file,
                file_name=file_path,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
    else:
        st.error("Please enter both the topic and content.") 
