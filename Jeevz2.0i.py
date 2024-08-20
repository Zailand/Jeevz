import streamlit as st
import os
import json
from importnb import Notebook
from pptx import Presentation
import pandas as pd
import tempfile

# Function to load an existing presentation
def load_presentation(file):
    presentation = Presentation(file)
    st.success(f"Presentation '{file.name}' has been successfully loaded.")
    return presentation

# Function to load shared data from the notes of the first slide
def load_shared_data(presentation):
    first_slide = presentation.slides[0]
    notes_slide = first_slide.notes_slide
    notes_text_frame = notes_slide.notes_text_frame
    notes_text = notes_text_frame.text

    if notes_text.strip():
        return json.loads(notes_text)
    else:
        return {}

# Function to ask the user where they would like to continue from
def continue_from():
    choice = st.selectbox(
        "Where would you like to continue from?",
        ["Hypothesis, Rationale & expected results", "Processing", "Compression conditions", "Tablet disintegration"]
    )
    return ["Hypothesis, Rationale & expected results", "Processing", "Compression conditions", "Tablet disintegration"].index(choice) + 1

# Function to prompt for continuation
def continue_prompt(step):
    with st.form(key=f"form_{step}"):
        col1, col2, col3 = st.columns([1, 0.1, 1])
        with col1:
            if step == 0:
                continue_button = st.form_submit_button("Continue to Hypothesis slide")
            elif step == 1:
                continue_button = st.form_submit_button("Continue to Process slide")
            elif step == 2:
                continue_button = st.form_submit_button("Continue to Compression conditions slide")
            elif step == 3:
                continue_button = st.form_submit_button("Continue to Disintegration conditions slide")
        with col2:
            st.write("or")
        with col3:
            download_button = st.form_submit_button("Download presentation")
        return continue_button, download_button

# Function to handle the download button click
def download_presentation(presentation):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        presentation.save(tmp.name)
        with open(tmp.name, "rb") as file:
            st.download_button(
                label="Download presentation",
                data=file,
                file_name="presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                key="download_button"
            )

# Function to check if a slide has been saved
def is_slide_saved(step):
    return st.session_state.get('saved_slides', {}).get(step, False)

# Function to mark a slide as saved
def mark_slide_as_saved(step):
    if 'saved_slides' not in st.session_state:
        st.session_state.saved_slides = {}
    st.session_state.saved_slides[step] = True

# Import functions from Functions.ipynb
with Notebook():
    from Functions import (
        title_slide,
        hypothesis_rationale_expected_slide,
        processing_slide,
        compression_conditions_slide,
        tablet_disintegration_slide,
    )

# Function to collect user inputs and store them temporarily for an existing project
def collect_user_inputs(presentation, shared_data, start_from=1):
    if start_from <= 1:
        st.write("Now working on the Hypothesis, Rationale & expected results slide")
        if not is_slide_saved(1):
            hypothesis_rationale_expected_slide(presentation, "temp.pptx", shared_data)
            mark_slide_as_saved(1)
        continue_button, download_button = continue_prompt(1)
        if continue_button:
            st.session_state.current_step = 2
        elif download_button:
            download_presentation(presentation)
        else:
            return False

    if start_from <= 2:
        st.write("Now working on the Processing slide")
        if not is_slide_saved(2):
            processing_slide(presentation, "temp.pptx", shared_data)
            mark_slide_as_saved(2)
        continue_button, download_button = continue_prompt(2)
        if continue_button:
            st.session_state.current_step = 3
        elif download_button:
            download_presentation(presentation)
        else:
            return False

    if start_from <= 3:
        st.write("Now working on the Compression conditions slide")
        if not is_slide_saved(3):
            compression_conditions_slide(presentation, "temp.pptx", shared_data)
            mark_slide_as_saved(3)
        continue_button, download_button = continue_prompt(3)
        if continue_button:
            st.session_state.current_step = 4
        elif download_button:
            download_presentation(presentation)
        else:
            return False

    if start_from <= 4:
        st.write("Now working on the Tablet disintegration slide")
        if not is_slide_saved(4):
            tablet_disintegration_slide(presentation, "temp.pptx", shared_data)
            mark_slide_as_saved(4)
        continue_button, download_button = continue_prompt(4)
        if continue_button:
            st.session_state.current_step = 5
        elif download_button:
            download_presentation(presentation)
        else:
            return False

    return True

# Function to collect user inputs and store them temporarily for a new project
def collect_user_inputs_new_project(presentation, shared_data):
    st.write("Now working on the Title Slide")
    if not is_slide_saved(0):
        title_slide(presentation, "temp.pptx", shared_data)
        mark_slide_as_saved(0)
    continue_button, download_button = continue_prompt(0)
    if continue_button:
        st.session_state.current_step = 1
    elif download_button:
        download_presentation(presentation)
    else:
        return False

    if not collect_user_inputs(presentation, shared_data, start_from=1):
        return False

    return True

# Function to start a new project
def start_new_project():
    st.write("Starting a new project...")
    presentation = Presentation()
    shared_data = {}

    if 'current_step' not in st.session_state:
        st.session_state.current_step = 0

    if st.session_state.current_step == 0:
        if collect_user_inputs_new_project(presentation, shared_data):
            pass
    else:
        if collect_user_inputs(presentation, shared_data, start_from=st.session_state.current_step):
            pass

# Function to load an existing project
def load_existing_project():
    st.write("Loading an existing project...")
    uploaded_file = st.file_uploader("Choose a PowerPoint file", type="pptx")
    if uploaded_file is not None:
        presentation = load_presentation(uploaded_file)
        shared_data = load_shared_data(presentation)
        start_from = continue_from()

        if 'current_step' not in st.session_state:
            st.session_state.current_step = start_from

        if collect_user_inputs(presentation, shared_data, start_from=st.session_state.current_step):
            pass

# Main function to ask the user if they want to start a new project or load an existing one
def main():
    st.title("PowerPoint Presentation Manager")
    choice = st.radio(
        "Would you like to:",
        ("Start a new project", "Load an existing project")
    )

    if choice == "Start a new project":
        start_new_project()
    elif choice == "Load an existing project":
        load_existing_project()

# Run the main function
if __name__ == "__main__":
    main()
