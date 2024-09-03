import base64
import pandas as pd
import streamlit as st
import time
from pathlib import Path
import base64
import random

# Set the page configuration for a wide layout
st.set_page_config(layout="wide")

# Custom CSS to center the heading, adjust table width, and center the Ganesha image
st.markdown(
    """
    <style>
    .center-heading {
        text-align: center;
    }
    .wide-table {
        width: 100%;
    }
    .ganesha-image {
        display: block;
        margin-left: auto;
        margin-right: auto;
        width: 150px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Streamlit UI
# Add the Ganesha image using Streamlit's image display function
image_path = Path("ganesha.jpg")
if image_path.exists():
    st.markdown(
        f'<div class="ganesha-image"><img src="data:image/jpg;base64,{base64.b64encode(open(str(image_path), "rb").read()).decode()}" alt="Ganesha"></div>',
        unsafe_allow_html=True
    )
else:
    st.error(f"Image file not found at {image_path.absolute()}. Please check the file name and location.")

st.markdown(
    "<h1 class='center-heading'>Ganesh Chaturthi Puja Slots Selection</h1>",
    unsafe_allow_html=True,
)

# Load the data from Excel file
@st.cache_data
def load_data():
    try:
        return pd.read_excel("names.xlsx")
    except FileNotFoundError:
        st.error("Error: The file 'names.xlsx' was not found. Please make sure it's in the same directory as this script.")
        return None
    except Exception as e:
        st.error(f"An error occurred while reading the file: {str(e)}")
        return None

df = load_data()

if df is not None:
    # Center the selection button and selected people list
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("Randomize Puja Slots"):
            # Show progress bar
            progress_bar = st.progress(0)
            for i in range(100):
                time.sleep(0.05)  # Sleep for 50ms
                progress_bar.progress(i + 1)
            
            # Randomize only the slots
            slots = df['Slots'].tolist()
            random.shuffle(slots)
            
            # Create a new DataFrame with randomized slots
            df_randomized = df[['Name', 'Flat no']].copy()
            df_randomized['Slots'] = slots
            
            # Sort the DataFrame based on slot timings
            df_sorted = df_randomized.sort_values('Slots')
            
            st.write("### Randomized and Sorted Puja Slots")
            # Increase the table width and display without indexes
            st.markdown(
                df_sorted.style.set_table_attributes('class="wide-table"').to_html(index=False),
                unsafe_allow_html=True,
            )
            
            # Provide download link for the randomized and sorted data
            csv = df_sorted.to_csv(index=False)
            st.download_button(label="Download Randomized and Sorted Slots as CSV", data=csv, file_name='randomized_sorted_puja_slots.csv', mime='text/csv')
        else:
            st.write("Click the button to randomly assign and sort Puja slots.")
