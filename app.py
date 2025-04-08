import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import os
from datetime import datetime

# Set up page configuration
st.set_page_config(page_title="Budget Performance Review", layout="wide")

# Title of the app
st.title("🧾 Y2024 Performance Review and Y2025 Outlook")

# Select pillar (objective)
pillar = st.selectbox("Select Pillar", ["Effective Governance", "Human Centric City", "Modern Infrastructure", "Thriving Economy"])

# Input section for project breakdown
total_projects = st.number_input("Total Projects Tracked", min_value=0)
not_started = st.number_input("Projects Not Yet Started", min_value=0)
in_progress = st.number_input("Projects In Progress", min_value=0)
completed = st.number_input("Projects Completed", min_value=0)

# Input for project status breakdown (by percentage)
green_projects = st.number_input("Projects (80-100% Complete)", min_value=0)
amber_projects = st.number_input("Projects (60-79% Complete)", min_value=0)
red_projects = st.number_input("Projects (0-59% Complete)", min_value=0)

# Calculate not completed projects
not_completed = total_projects - completed

# Show results when button is clicked
if st.button("Generate Summary"):
    # Display summary information
    st.subheader(f"Performance Summary for {pillar}")
    st.write(f"- Total Projects: {total_projects}")
    st.write(f"- Not Started: {not_started}")
    st.write(f"- In Progress: {in_progress}")
    st.write(f"- Completed: {completed}")
    st.write(f"- Not Completed: {not_completed}")
    st.write(f"- Projects with Green Status (80-100%): {green_projects}")
    st.write(f"- Projects with Amber Status (60-79%): {amber_projects}")
    st.write(f"- Projects with Red Status (0-59%): {red_projects}")

    # Expert Analysis based on the project status breakdown
    if green_projects >= amber_projects and green_projects >= red_projects:
        st.write("🔍 Expert Analysis: The majority of the projects are progressing well, with green status indicating timely completion. However, attention should still be given to the red and amber projects to ensure any delays are managed effectively.")
    elif amber_projects >= green_projects and amber_projects >= red_projects:
        st.write("🔍 Expert Analysis: A significant portion of the projects are in the amber status, meaning there are slight delays. Focus on streamlining processes to avoid further delays and get these projects back on track.")
    else:
        st.write("🔍 Expert Analysis: The red status indicates serious delays across multiple projects. Immediate intervention is required to understand the bottlenecks and correct course to avoid jeopardizing the objective for 2025.")

    # 2025 Outlook (General suggestions)
    st.subheader("2025 Outlook and Suggestions")
    if total_projects > 0:
        st.write("📅 The projects that are delayed need immediate attention to get back on schedule. Prioritizing green status projects for timely completion is key to achieving the pillar's goals.")
        st.write("💡 Consider allocating more resources to the red and amber projects for faster progress and ensure all projects meet their deadlines for a successful 2025.")

# Save results as an image or ppt
def save_slide_and_image():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])

    # Add Title
    title = slide.shapes.title
    title.text = f"Performance Summary for {pillar}"

    # Add Text
    txBox = slide.shapes.add_textbox(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.5), Inches(9), Inches(5.5))
    tf = txBox.text_frame
    tf.text = f"""
    Total Projects: {total_projects}
    Not Started: {not_started}
    In Progress: {in_progress}
    Completed: {completed}
    Not Completed: {not_completed}
    Green Projects (80-100%): {green_projects}
    Amber Projects (60-79%): {amber_projects}
    Red Projects (0-59%): {red_projects}
    """

    # Save as ppt
    output_ppt = os.path.join("output_slides", f"Performance_Summary_{pillar}.pptx")
    prs.save(output_ppt)

    # Optionally save as PNG (rendering the slide as image)
    slide_image_path = os.path.join("output_slides", f"Performance_Summary_{pillar}.png")
    slide.shapes[0].element.getparent().getparent().getparent().save(slide_image_path)

    return output_ppt, slide_image_path

# Button to save results as ppt and image
if st.button("Save as PPT and Image"):
    ppt_file, image_file = save_slide_and_image()
    st.write(f"✅ Summary saved as PPT: {ppt_file}")
    st.write(f"✅ Summary saved as Image: {image_file}")

