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
    The output JSON must completely reflect the user instructions.
    Expected output format:
    
    {
      "slides": [
        {
          "title": "Slide Title",
          "content": "Full text content rewritten to reflect the user instructions.",
          "keywords": ["keyword1", "keyword2"]
        },
        ...
      ]
    }
    
    The output should contain only the JSON.
    """
    prompt = f"""
You are an assistant that completely rewrites presentation content based solely on the user instructions provided.
Do not include any default template text.
Your output should be a JSON object with a single key "slides" containing an array of slide objects.
Each slide object must have:
 - "title": a concise slide title.
 - "content": full text content that completely reflects the user instructions.
 - "keywords": an array of keywords for image selection.

User instructions: {user_instructions}

Please output only the JSON with no additional commentary or markdown formatting.
Example output:
{{
  "slides": [
    {{
      "title": "Introduction",
      "content": "This slide introduces the topic with rewritten content based on the user instructions.",
      "keywords": ["introduction", "overview"]
    }},
    {{
      "title": "Details",
      "content": "This slide provides detailed, rewritten content reflecting the user input.",
      "keywords": ["details", "information"]
    }}
  ]
}}
"""
    try:
        response = openai.chat.completions.create(
            model="gpt-4o",  # Using GPT-4o as requested
            messages=[
                {"role": "system", "content": "You generate presentation slide content in JSON based solely on user instructions."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
        )
        response_content = response.choices[0].message.content.strip()
        st.write("Raw GPT Output:", response_content)  # Debug: view raw output
        
        # Remove markdown code fences if present
        if response_content.startswith("```"):
            lines = response_content.splitlines()
            if len(lines) > 2:
                response_content = "\n".join(lines[1:-1]).strip()
            else:
                response_content = response_content.strip("```").strip()
        
        slides_data = json.loads(response_content)
        return slides_data.get("slides", [])
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

def remove_all_slides(prs):
    """
    Completely remove all slides from the presentation.
    This uses a while loop on the underlying XML element to ensure
    that every slide is removed.
    """
    sldIdLst = prs.slides._sldIdLst
    while len(sldIdLst) > 0:
        sldIdLst.remove(sldIdLst[0])

def create_presentation(slides, template_path):
    """
    Create a PowerPoint presentation using the provided slide data and a template.
    The template is used only for its theme (colors, fonts, layout, etc.).
    All original slides are removed so that the output presentation includes only
    the content generated based on user instructions.
    """
    try:
        prs = Presentation(template_path)
    except Exception as e:
        st.error(f"Failed to load template: {e}")
        return None

    # Remove all slides from the template so that we start with a blank slate.
    remove_all_slides(prs)
    st.write("Slide count after removal:", len(prs.slides))

    for slide_data in slides:
        # Use layout index 1 (commonly "Title and Content")
        try:
            slide_layout = prs.slide_layouts[1]
        except IndexError:
            st.error("The selected template does not have the expected slide layout.")
            continue

        slide = prs.slides.add_slide(slide_layout)

        # Set the slide title.
        try:
            title_placeholder = slide.shapes.title
            if title_placeholder:
                title_placeholder.text = slide_data.get("title", "")
        except Exception:
            st.warning("This slide does not have a title placeholder.")

        # Set the slide content. If the content placeholder is missing, add a textbox.
        content_set = False
        try:
            content_placeholder = slide.placeholders[1]
            content_placeholder.text = slide_data.get("content", "")
            content_set = True
        except (IndexError, KeyError):
            pass

        if not content_set:
            left = Inches(1)
            top = Inches(2)
            width = Inches(8)
            height = Inches(3)
            textbox = slide.shapes.add_textbox(left, top, width, height)
            textbox.text_frame.text = slide_data.get("content", "")

        # Add an image if a matching one is found.
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
    st.write(
        "Enter your presentation instructions below. The output presentation will completely reflect your input text. "
        "A selected template is used only for styling (colors, fonts, layout), while all text and images are generated "
        "based on your instructions."
    )

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
        
        if slides is None or len(slides) == 0:
            st.error("Failed to generate slide content. Check your instructions and try again.")
            return
        
        st.subheader("Generated Slide Data")
        st.json(slides)
        
        with st.spinner("Creating PowerPoint presentation..."):
            prs = create_presentation(slides, template_path)
            if prs is None:
                st.error("Error creating presentation.")
                return

            # Save the presentation to a BytesIO object.
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
