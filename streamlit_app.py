import streamlit as st

from st_files_connection import FilesConnection

st.title("🎈 My new app")
st.write(
    "Let's start building! For help and inspiration, head over to [docs.streamlit.io](https://docs.streamlit.io/)."
)
st.write(
    "Hello World!"
)

from smart_open import smart_open

#AWS Connection
#aws_key=os.environ['AWS_ACCESS_KEY_ID']
#aws_secret=os.environ['AWS_SECRET_ACCESS_KEY']
bucket_name = 'h4-hack-week-aug-2024'
object_key = 'myfile.csv'
path = 's3://{}:{}@{}/{}'.format("AKIA5DRNUTKHYRXM5LPV", "KOU2S7VqQQmB4974Vl5Ve0CRMxBPZ55RwR0HFM1O", bucket_name, object_key)


#Connecting to AWS through smart_open python package and getting the data
@st.experimental_memo
def load_data(path):
    data = pd.read_csv(smart_open(path),index_col=0)
    return data

df = load_data(path) #saving the returned data into dataframe
st.write(df)
