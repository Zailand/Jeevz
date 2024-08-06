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

# Function to prompt for continuation with a dropdown menu
def continue_prompt_with_choice(current_step, choices):
    next_step = st.selectbox("Continue to", choices, index=0)
    if st.button("Continue"):
        return next_step
    return None

# Import functions from Functions.ipynb
with Notebook():
    from Functions import (
        title_slide,
        hypothesis_rationale_expected_slide,
        processing_slide,
        compression_conditions_slide,
        tablet_disintegration_slide,
    )

# Function to handle the progression through the steps
def handle_steps(presentation, presentation_path, shared_data):
    if 'current_step' not in st.session_state:
        st.session_state['current_step'] = 0

    current_step = st.session_state['current_step']
    st.write(f"Current step: {current_step}")

    steps = [
        ("Title Slide", title_slide),
        ("Hypothesis, Rationale & Expected Results Slide", hypothesis_rationale_expected_slide),
        ("Processing Slide", processing_slide),
        ("Compression Conditions Slide", compression_conditions_slide),
        ("Tablet Disintegration Slide", tablet_disintegration_slide)
    ]

    if current_step < len(steps):
        step_name, step_function = steps[current_step]
        st.write(f"Now working on the {step_name}")
        step_function(presentation, presentation_path, shared_data)

        # Determine the available choices for the next step
        choices = [f"{i}. {steps[i][0]}" for i in range(current_step + 1, len(steps))]
        next_step = continue_prompt_with_choice(current_step, choices)

        if next_step is not None:
            st.session_state['current_step'] = int(next_step.split(".")[0])
            st.experimental_rerun()  # Ensure the UI updates
    else:
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
    presentation = Presentation()
    shared_data = {}

    handle_steps(presentation, presentation_path, shared_data)

# Function to load an existing project
def load_existing_project():
    st.write("Loading an existing project...")
    uploaded_file = st.file_uploader("Choose a PowerPoint file", type="pptx")
    if uploaded_file is not None:
        presentation = load_presentation(uploaded_file)
        shared_data = load_shared_data(presentation)
        handle_steps(presentation, uploaded_file.name, shared_data)

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
