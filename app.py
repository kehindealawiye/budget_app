import streamlit as st
import matplotlib.pyplot as plt
import os
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from datetime import datetime

# Create output folder if it doesn't exist
if not os.path.exists("output_slides"):
    os.makedirs("output_slides")

st.set_page_config(page_title="Budget Pillars Performance Review", layout="centered")

# Page styling
st.markdown("""
    <style>
        body {
            background-color: #0b1e3f;
            color: white;
        }
        .stApp {
            background-color: #0b1e3f;
        }
        h1, h2, h3, h4, h5, h6 {
            color: white;
        }
    </style>
""", unsafe_allow_html=True)

st.title("Assessing Y2024 Performance and Charting the Course for Y2025")

# Pillar input
pillar = st.selectbox("Select Pillar", [
    "Effective Governance",
    "Human Centric City",
    "Modern Infrastructure",
    "Thriving Economy"
])

st.markdown("### Project Status Input")
total_projects = st.number_input("Total number of projects tracked", min_value=0)
not_commenced = st.number_input("Projects yet to commence", min_value=0, max_value=total_projects)
initiation = st.number_input("Projects at initiation", min_value=0, max_value=total_projects)
in_progress = st.number_input("Projects in progress", min_value=0, max_value=total
