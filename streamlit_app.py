import streamlit as st
import pandas as pd

from st_files_connection import FilesConnection

st.title("🎈 My new app")
st.write(
    "Let's start building! For help and inspiration, head over to [docs.streamlit.io](https://docs.streamlit.io/)."
)
st.write(
    "Hello World!"
)

from st_files_connection import FilesConnection

# Create connection object and retrieve file contents.
# Specify input format is a csv and to cache the result for 600 seconds.
conn = st.connection('s3', type=FilesConnection)
st.write("conn obtained")
df = conn.read("fbc-hackathon-test/Operations ScoreCard - UX.csv", input_format="csv", ttl=600)
st.write("df obtained")
st.table(df)
# Print results.
#for row in df.itertuples():
    #st.write(f"{row}")


