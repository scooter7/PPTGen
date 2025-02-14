import streamlit as st
import openai
import json
import os
import difflib
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO

# Set OpenAI API key from Streamlit secrets
openai.api_key = st.secrets["openai_api_key"]

def get_template_file():
    """
    Let the user select a template from the "templates" folder.
    The chosen file (e.g. powerpointtemplate.pptx) should contain a single slide
    that defines the base design.
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
    Use GPT-4o to generate a JSON object containing slide data based solely on the
    user instructions. Expected JSON format:
    
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
    
    Output only the JSON.
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
        st.write("Raw GPT Output:", response_content)  # Debug output
        
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

def select_best_image_for_slide(slide_data, images_folder="images", threshold=0.1):
    """
    Given a slide's data, this function computes a similarity score between the slide's
    combined text (content + keywords) and each image filename in the images folder.
    The image with the highest score (above the given threshold) is returned as the best fit.
    """
    if not os.path.isdir(images_folder):
        return None

    # Combine slide content and keywords for matching.
    slide_text = slide_data.get("content", "") + " " + " ".join(slide_data.get("keywords", []))
    slide_text = slide_text.lower()

    best_image = None
    best_ratio = 0

    for filename in os.listdir(images_folder):
        if filename.lower().endswith((".jpg", ".jpeg", ".png")):
            ratio = difflib.SequenceMatcher(None, slide_text, filename.lower()).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_image = os.path.join(images_folder, filename)
    
    if best_ratio < threshold:
        return None
    return best_image

def create_presentation(slides, template_path):
    """
    Create a new presentation using the provided template as the base design.
    The template is assumed to have a single slide; this slide is updated with the first
    slide's content, and additional slides are created using the same slide layout.
    Text wrapping is enabled so that long text is properly wrapped.
    Images are inserted based on best-fit matching with the slide's text.
    """
    prs = Presentation(template_path)
    if len(prs.slides) == 0:
        st.error("Template does not contain any slides.")
        return None

    # Use the base slide from the template.
    base_slide = prs.slides[0]
    first_slide_data = slides[0]

    # Update the base slide's title.
    if base_slide.shapes.title:
        base_slide.shapes.title.text = first_slide_data.get("title", "")
        base_slide.shapes.title.text_frame.word_wrap = True

    # Update the base slide's content.
    try:
        content_placeholder = base_slide.placeholders[1]
        content_placeholder.text = first_slide_data.get("content", "")
        content_placeholder.text_frame.word_wrap = True
    except (IndexError, KeyError):
        left = Inches(1)
        top = Inches(2)
        width = Inches(8)
        height = Inches(3)
        textbox = base_slide.shapes.add_textbox(left, top, width, height)
        textbox.text_frame.text = first_slide_data.get("content", "")
        textbox.text_frame.word_wrap = True

    # Insert an image on the base slide based on best-fit matching.
    best_image = select_best_image_for_slide(first_slide_data)
    if best_image and os.path.exists(best_image):
        base_slide.shapes.add_picture(best_image, Inches(5), Inches(1), height=Inches(4))

    # Retrieve the base slide's layout for creating new slides.
    base_layout = base_slide.slide_layout

    # Create additional slides.
    for slide_data in slides[1:]:
        new_slide = prs.slides.add_slide(base_layout)
        if new_slide.shapes.title:
            new_slide.shapes.title.text = slide_data.get("title", "")
            new_slide.shapes.title.text_frame.word_wrap = True
        try:
            content_placeholder = new_slide.placeholders[1]
            content_placeholder.text = slide_data.get("content", "")
            content_placeholder.text_frame.word_wrap = True
        except (IndexError, KeyError):
            left = Inches(1)
            top = Inches(2)
            width = Inches(8)
            height = Inches(3)
            textbox = new_slide.shapes.add_textbox(left, top, width, height)
            textbox.text_frame.text = slide_data.get("content", "")
            textbox.text_frame.word_wrap = True

        # Insert the best-fit image for this slide.
        best_image = select_best_image_for_slide(slide_data)
        if best_image and os.path.exists(best_image):
            new_slide.shapes.add_picture(best_image, Inches(5), Inches(1), height=Inches(4))
    return prs

def main():
    st.title("AI-Powered PowerPoint Presentation Generator")
    st.write(
        "Enter your presentation instructions below. The output presentation will fully reflect your input text, "
        "using the selected template's base slide design. Text will wrap appropriately, and images will be inserted "
        "based on matching the slide's text to the best-fit image from the images folder."
    )

    # Use the template file from the templates folder.
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
