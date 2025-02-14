import streamlit as st
import openai
import json
import os
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO

# Set OpenAI API key from Streamlit secrets
openai.api_key = st.secrets["openai_api_key"]

def get_template_file():
    """
    Lists available template files from the "templates" folder and
    lets the user select one. Supports both .pptx and .ppt files.
    """
    templates_folder = "templates"
    if not os.path.isdir(templates_folder):
        st.error("Templates folder not found!")
        return None

    template_files = [f for f in os.listdir(templates_folder) if f.lower().endswith(('.pptx', '.ppt'))]
    if not template_files:
        st.error("No template files found in the templates folder!")
        return None

    selected_template = st.selectbox("Select a Template", template_files)
    template_path = os.path.join(templates_folder, selected_template)
    if selected_template.lower().endswith('.ppt'):
        st.warning("PPT templates are not directly supported. Please use a PPTX template or convert your PPT to PPTX.")
    return template_path

def generate_slides(user_instructions):
    """
    Use OpenAI ChatCompletion to generate slide content using GPT-4o.
    We expect a JSON output of the following form:
    {
      "slides": [
        {
          "title": "Slide Title",
          "content": "Slide content as bullet points or a paragraph.",
          "keywords": ["keyword1", "keyword2"]
        },
        ...
      ]
    }
    """
    prompt = f"""
You are an assistant that creates slide content for a presentation.
Based on the following instructions, generate a JSON with an array of slides.
User instructions: {user_instructions}
Please output valid JSON in the following format:

{{
  "slides": [
    {{
      "title": "Slide Title",
      "content": "Slide content for the slide. It may be bullet points or a short paragraph.",
      "keywords": ["keyword1", "keyword2"]
    }},
    ...
  ]
}}

Only output the JSON and nothing else.
"""
    try:
        response = openai.chat.completions.create(
            model="gpt-4o",  # Using GPT-4o as requested
            messages=[
                {"role": "system", "content": "You generate presentation slide content in JSON."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
        )
        response_content = response.choices[0].message.content.strip()
        st.write("Raw output from GPT:", response_content)  # Debug line to check the output
        
        # Remove markdown code block markers if present
        if response_content.startswith("```"):
            # Remove the first line (```json) and the last line (```)
            lines = response_content.splitlines()
            # If there are more than 2 lines, join them excluding the first and last
            if len(lines) > 2:
                response_content = "\n".join(lines[1:-1]).strip()
            else:
                response_content = response_content.strip("```").strip()
        
        slides_data = json.loads(response_content)
        return slides_data["slides"]
    except Exception as e:
        st.error(f"Error generating slides: {e}")
        return None

def select_image_for_slide(keywords, images_folder="images"):
    """
    Given a list of keywords, search for an image in the images_folder 
    whose filename contains one of the keywords. Only considers jpg, jpeg, and png.
    """
    if not os.path.isdir(images_folder):
        return None

    for filename in os.listdir(images_folder):
        if filename.lower().endswith((".jpg", ".jpeg", ".png")):
            for kw in keywords:
                if kw.lower() in filename.lower():
                    return os.path.join(images_folder, filename)
    return None

def create_presentation(slides, template_path):
    """
    Create a PowerPoint presentation using the provided slide data and a template.
    Note: python-pptx supports only PPTX templates. If a PPT file is selected,
    this function may fail.
    """
    try:
        prs = Presentation(template_path)
    except Exception as e:
        st.error(f"Failed to load template: {e}")
        return None

    for slide_data in slides:
        # Attempt to use layout index 1 (commonly "Title and Content")
        try:
            slide_layout = prs.slide_layouts[1]
        except IndexError:
            st.error("The selected template does not have the expected slide layout.")
            continue

        slide = prs.slides.add_slide(slide_layout)

        # Set the slide title if the title placeholder exists.
        if hasattr(slide.shapes, "title") and slide.shapes.title is not None:
            title_placeholder = slide.shapes.title
            title_placeholder.text = slide_data.get("title", "")
        else:
            st.warning("This slide does not have a title placeholder.")

        # Set the slide content if a content placeholder exists.
        try:
            content_placeholder = slide.placeholders[1]
            content_placeholder.text = slide_data.get("content", "")
        except IndexError:
            st.warning("This slide does not have a content placeholder.")

        # Try to select an image based on the slide keywords.
        keywords = slide_data.get("keywords", [])
        image_path = select_image_for_slide(keywords)
        if image_path and os.path.exists(image_path):
            left = Inches(5)
            top = Inches(1)
            height = Inches(4)
            slide.shapes.add_picture(image_path, left, top, height=height)

    return prs

def main():
    st.title("AI-Powered PowerPoint Presentation Generator")
    st.write("Enter your presentation instructions below and generate a customized PowerPoint presentation based on your selected template and images.")

    # Let the user select a template file from the "templates" folder.
    template_path = get_template_file()
    if template_path is None:
        st.stop()

    user_instructions = st.text_area("Presentation Instructions", height=150)

    if st.button("Generate Presentation"):
        if not user_instructions.strip():
            st.warning("Please enter your presentation instructions.")
            return
        
        with st.spinner("Generating slide content..."):
            slides = generate_slides(user_instructions)
        
        if slides is None:
            st.error("Failed to generate slide content.")
            return
        
        st.subheader("Generated Slide Data")
        st.json(slides)
        
        with st.spinner("Creating PowerPoint presentation..."):
            prs = create_presentation(slides, template_path)
            if prs is None:
                st.error("Error creating presentation.")
                return

            # Save the presentation to a BytesIO object
            pptx_io = BytesIO()
            prs.save(pptx_io)
            pptx_io.seek(0)
        
        st.success("Presentation created successfully!")
        st.download_button(
            label="Download Presentation",
            data=pptx_io,
            file_name="generated_presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

if __name__ == "__main__":
    main()
