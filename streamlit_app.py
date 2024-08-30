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
import io
import csv
import ast
from datetime import date
import requests
import base64
import plotly.express as px

# Load environment variables from a .env file if present
load_dotenv()
path = os.path.dirname(__file__)
# Set your OpenAI API key
openapi_key = st.text_input("OpenAI API Key", type="password")
openai.api_key = openapi_key
today = date.today()

# Title of the Streamlit app
st.title('Slide Content Generator')
def get_franchise_data(topic):
    conn = st.connection('s3', type=FilesConnection)
    df = conn.read("fbc-hackathon-test/growth.csv", input_format="csv", ttl=600) 
    df = df.transpose()
    df.columns = df.iloc[0]  # Use the first row as the header
    df = df.drop(df.index[0])  # Drop the first row since it is now the header
    df = df.to_dict()
    keys_to_keep = topic.split(',')
    keys_to_keep = [key.strip() for key in topic.split(',')]
    filtered_dict = {}
    for data in df:
        if str(data) in keys_to_keep:
            filtered_dict[str(data)] = df[data]

    st.write(filtered_dict)
    return filtered_dict

def get_median_data():
    conn = st.connection('s3', type=FilesConnection)
    df = conn.read("fbc-hackathon-test/Network_Median.csv", input_format="csv", ttl=600) 
    df = df.drop(df.index[0])  # Drop the first row since it is now the header
    df = df.to_dict()
    st.write("median data")
    st.write(df)
    return df

def add_image(slide, image, left, top, width):
    """function to add an image to the PowerPoint slide and specify its position and width"""
    slide.shapes.add_picture(image, left=left, top=top, width=width)
    
# Function to generate slide content
def generate_aggregate_metrics(content):
    prompt_txt = f"Wait for user input to return a response. Use this data to generate the output as a valid dictionary object:\n\n{str(content)}"
    prompt = f"You are a helpful assistant that generates an executive summary of Franchise's performance metrics. Calculate aggregate metrics for given Franchises and return output a valid dictionary object with key as aggregate_metrics. Do not return anything else."

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
    return ast.literal_eval(generated_text)

def generate_key_insights(content):
    prompt_txt = f"Wait for user input to return a response. Use this data to generate the output as a valid dictionary object:\n\n{str(content)}"
    prompt = f"You are a helpful assistant that generates an executive summary of Franchise's performance metrics. Analyze the data and summarize the following trends as brief bullets comparing the performance of the franchises in the enterprise. If the franchise has shown growth, what are the factors contributing to it according to the data given?"
   
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
    #st.write(f"Response: {generated_text}")
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



def generate_graph(aggregate_metrics):

    # Step 1: Generate the graph image using OpenAPI API
    # (Example placeholder code - this will depend on your specific API)

    api_url = "https://api.openai.com/v1/images/generate"
    headers = {"Authorization": "Bearer " + openapi_key}
    data = {
        "prompt": "A bar graph showing total revenue growth between current and previous years from this data: {aggregate_metrics}",
        "size": "1024x1024"
    }
    response = requests.post(api_url, headers=headers, json=data)
    st.write(response)
    st.write(response.json())
    # Assuming the API returns an image in base64 format
    image_data = base64.b64decode(response.json()['image'])

    # Step 2: Save the image to a file
    with open("graph.png", "wb") as image_file:
        image_file.write(image_data)

    st.write('image data received')

# Function to create a PowerPoint presentation
def create_presentation(franchise_data, slide_content, key_insights):
    pptx = path + '//' + 'template.pptx'
    prs = Presentation(pptx)
    bullet_slide_layout = prs.slide_layouts.get_by_name('Purple_Circle_Corners')

    details_dict = {
    'Franchisee': 'Franchisee',
    'NetworkPerformancePartner': 'FBC',
    'Region': 'DO',
    'WeightedScore': 'Your Total Score',
    'Rank': 'Network Ranking',
    'CurrentYearTotalBillableHours': 'Current Year Billable Hours',
    'LastYearTotalBillableHours' : 'Previous Year Billable Hours', 
    'CurrentYearTotalRevenue': 'Current Year Total Revenue',
    'LastYearYearTotalRevenue' : 'Previous Year Total Revenue',     
    'HoursGrowth': 'Your Hours Growth %',
    'RevenueGrowth': 'Your Revenue Growth %',
    'HoursGrowth': 'Your Hours Growth %', 
    'RevenueGrowth': 'Your Revenue Growth %'
    }

    median_data = get_median_data()

    owner = []
    owner_full_name = []
    aggregate_metrics = {}
    #key_insights = {}
    #st.write(f"so far so good")
    #st.write(isinstance(slide_content, dict))
    for item in slide_content:
        st.write(item)
    if 'aggregate_metrics' in slide_content:
        aggregate_metrics = slide_content['aggregate_metrics']
    if 'AggregateMetrics' in slide_content:
        aggregate_metrics = slide_content['AggregateMetrics']
    #st.write(aggregate_metrics)

    franchise_numbers = []
    
    for franchise in franchise_data:
        franchise_numbers.append(str(franchise))
        slide = prs.slides.add_slide(bullet_slide_layout)
        shapes = slide.shapes
        title_shape = shapes.title
        body_shape = shapes.placeholders[1]
        tf = body_shape.text_frame
        ind_fran = franchise_data[str(franchise)]
        owner_full_name.append(ind_fran['FirstName'] + ' ' + ind_fran['LastName']) 
        owner.append(ind_fran['LastName'])
        title_shape.text = f"Franchise {franchise} - {ind_fran['FirstName']} {ind_fran['LastName']}"
        for k in ind_fran:
            if k in details_dict:
                p = tf.add_paragraph()
                p.text+= f"  {details_dict[k]}: {ind_fran[k]}\n\n"

    franchise_numbers_string = ", ".join(franchise_numbers)
    owner = list(set(owner))
    #st.write(owner)
    # Convert the array to a comma-separated string
    owner_string = ", ".join(owner)
    #st.write(f"comma separated unique list: {owner_string}")

    owner_full_name = list(set(owner_full_name))
    owner_full_name_string = ", ".join(owner_full_name)
    st.write(f"comma separated unique list: {owner_full_name_string}")
    
    first_slide = prs.slides[0]
    shapes_1 = []

    third_slide = prs.slides[2]
    shapes_2 = []

    #Aggregate Metrics
    slide = prs.slides.add_slide(bullet_slide_layout)
    shapes = slide.shapes
    title_shape = shapes.title
    title_shape.text = f"Enterprise Business Overview"
    body_shape = shapes.placeholders[1]
    tf = body_shape.text_frame
    for k in aggregate_metrics:
        p = tf.add_paragraph()
        p.text+= f"  {k}: {aggregate_metrics[k]}\n\n"

    df = pd.DataFrame(
        [["Current", 1234567, 3450, 100], ["Previous", 8758758, 73877, 800]],
        columns=["Year", "Revenue", "Billable Hours", "RPN Leads"]
    )

    fig = px.bar(df, x="Year", y=["Revenue", "Billable Hours", "RPN Leads"], barmode='group', height=400)
    # st.dataframe(df) # if need to display dataframe
    st.plotly_chart(fig)

    fig.write_image("metrics.png")
    metrics_im = 'metrics.png'

    add_image(prs.slides[4], image=metrics_im, left=left, width=width, top=top)
    os.remove('metrics.png')

    #Key Insights
    slide = prs.slides.add_slide(bullet_slide_layout)
    shapes = slide.shapes
    title_shape = shapes.title
    title_shape.text = f"Key Insights"
    body_shape = shapes.placeholders[1]
    tf = body_shape.text_frame
    p = tf.add_paragraph()
    p.text+= key_insights
        
    # create lists with shape objects
    for shape in first_slide.shapes:
        shapes_1.append(shape)

    # create lists with shape objects
    for shape in third_slide.shapes:
        shapes_2.append(shape)

    # initiate a dictionary of placeholders and values to replace
    replaces_1 = {
        '{owner}': owner_string,
        '{date}': today }
    replace_text(replaces_1, shapes_1)

    replaces_2 = {
        '{owner_name}': owner_full_name_string,
        '{franchise_numbers}': franchise_numbers_string }
    replace_text(replaces_2, shapes_2)

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
        st.write(franchise_data)
        generated_content = generate_aggregate_metrics(franchise_data)
        
        key_insights = generate_key_insights(franchise_data)
        st.subheader("Generated Slide Content:")
        
        # Create and offer download of the PowerPoint presentation
        file_path = create_presentation(franchise_data, generated_content, key_insights)
        
        with open(file_path, "rb") as file:
            btn = st.download_button(
                label="Download PowerPoint Presentation",
                data=file,
                file_name=file_path,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
    else:
        st.error("Please enter both the topic and content.") 
