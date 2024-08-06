import streamlit as st
import os
import json
from importnb import Notebook
from pptx import Presentation
import pandas as pd
import base64

# Function to load an existing presentation
def load_presentation(file):
    presentation = Presentation(file)
    st.session_state['presentation'] = presentation
    st.session_state['presentation_path'] = file.name
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

# Import functions from Functions.ipynb
with Notebook():
    from Functions import (
        title_slide,
        hypothesis_rationale_expected_slide,
        processing_slide,
        compression_conditions_slide,
        tablet_disintegration_slide,
    )

# Function to prompt for continuation
def continue_prompt():
    return st.button("Continue")

# Function to handle the progression through the steps
def handle_steps():
    if 'current_step' not in st.session_state:
        st.session_state['current_step'] = 0

    current_step = st.session_state['current_step']
    st.write(f"Current step: {current_step}")

    presentation = st.session_state.get('presentation')
    presentation_path = st.session_state.get('presentation_path')
    shared_data = st.session_state.get('shared_data', {})

    if current_step == 0:
        st.write("Now working on the Title Slide")
        title_slide(presentation, presentation_path, shared_data)
        if continue_prompt():
            save_presentation(presentation, presentation_path)
            st.session_state['current_step'] = 1
    elif current_step == 1:
        st.write("Now working on the Hypothesis, Rationale & expected results slide")
        hypothesis_rationale_expected_slide(presentation, presentation_path, shared_data)
        if continue_prompt():
            save_presentation(presentation, presentation_path)
            st.session_state['current_step'] = 2
    elif current_step == 2:
        st.write("Now working on the Processing slide")
        processing_slide(presentation, presentation_path, shared_data)
        if continue_prompt():
            save_presentation(presentation, presentation_path)
            st.session_state['current_step'] = 3
    elif current_step == 3:
        st.write("Now working on the Compression conditions slide")
        compression_conditions_slide(presentation, presentation_path, shared_data)
        if continue_prompt():
            save_presentation(presentation, presentation_path)
            st.session_state['current_step'] = 4
    elif current_step == 4:
        st.write("Now working on the Tablet disintegration slide")
        tablet_disintegration_slide(presentation, presentation_path, shared_data)

    st.write("or")
    if st.button("Download presentation"):
        if save_presentation(presentation, presentation_path):
            provide_download_link(presentation_path)

# Function to save the presentation with error handling
def save_presentation(presentation, presentation_path):
    try:
        presentation.save(presentation_path)
        st.success(f"Presentation saved as {presentation_path}")
        return True
    except IOError:
        st.error(f"The file '{presentation_path}' cannot be saved. It might be open or you might not have permission.")
        return False

# Function to provide a download link for the presentation
def provide_download_link(presentation_path):
    with open(presentation_path, "rb") as file:
        btn = st.download_button(
            label="Download Presentation",
            data=file,
            file_name=presentation_path,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    return btn

# Function to start a new project
def start_new_project():
    st.write("Starting a new project...")
    presentation_path = st.text_input("Enter the path to save the new presentation:", "new_presentation.pptx")
    if presentation_path:
        presentation = Presentation()
        st.session_state['presentation'] = presentation
        st.session_state['presentation_path'] = presentation_path
        st.session_state['shared_data'] = {}
        handle_steps()

# Function to load an existing project
def load_existing_project():
    st.write("Loading an existing project...")
    uploaded_file = st.file_uploader("Choose a PowerPoint file", type="pptx")
    if uploaded_file is not None:
        presentation = load_presentation(uploaded_file)
        shared_data = load_shared_data(presentation)
        st.session_state['shared_data'] = shared_data
        handle_steps()

# Main function to ask the user if they want to start a new project or load an existing one
def main():
    st.title("PowerPoint Presentation Manager")
    choice = st.radio(
        "Would you like to:",
        ("Start a new project", "Load an existing project")
    )

    if choice == "Start a new project":
        st.session_state['current_step'] = 0
        start_new_project()
    elif choice == "Load an existing project":
        st.session_state['current_step'] = 0
        load_existing_project()

# Run the main function
if __name__ == "__main__":
    main()
