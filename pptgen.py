import streamlit as st
import openai
import json
import os
import torch
from transformers import CLIPProcessor, CLIPModel
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
from PIL import Image

# Set OpenAI API key from Streamlit secrets
openai.api_key = st.secrets["openai_api_key"]

# Load CLIP model and processor (this may take a few seconds on first load)
device = "cuda" if torch.cuda.is_available() else "cpu"
@st.cache_resource
def load_clip_model():
    model = CLIPModel.from_pretrained("openai/clip-vit-base-patch32")
    processor = CLIPProcessor.from_pretrained("openai/clip-vit-base-patch32")
    model.to(device)
    return model, processor

clip_model, clip_processor = load_clip_model()

def get_content_template():
    """
    Return the path to the content template (powerpointtemplate.pptx) that defines the theme.
    """
    template_path = os.path.join("templates", "powerpointtemplate.pptx")
    if not os.path.exists(template_path):
        st.error("Content template not found in the templates folder!")
        return None
    return template_path

def generate_slides(user_instructions):
    """
    Use GPT-4o to generate slide content in JSON format.
    Expected output:
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
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You generate presentation slide content in JSON based solely on user instructions."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
        )
        response_content = response.choices[0].message.content.strip()
        st.write("Raw GPT Output:", response_content)  # Debug output

        # Remove markdown code fences if present.
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

def select_best_image_for_slide(slide_data, images_folder="images"):
    """
    Use CLIP to compute the similarity between the slide's text and each image.
    The slide's text is computed from its title, content, and keywords.
    Returns the image path with the highest cosine similarity.
    """
    if not os.path.isdir(images_folder):
        return None

    slide_text = " ".join([
        slide_data.get("title", ""),
        slide_data.get("content", ""),
        " ".join(slide_data.get("keywords", []))
    ]).strip()
    if not slide_text:
        return None

    text_inputs = clip_processor(text=[slide_text], return_tensors="pt", padding=True)
    with torch.no_grad():
        text_features = clip_model.get_text_features(**text_inputs).to(device)
    text_features = text_features / text_features.norm(dim=-1, keepdim=True)

    best_score = -1.0
    best_image = None

    for filename in os.listdir(images_folder):
        if filename.lower().endswith((".jpg", ".jpeg", ".png")):
            image_path = os.path.join(images_folder, filename)
            try:
                image = Image.open(image_path).convert("RGB")
            except Exception as e:
                st.write(f"Error opening image {filename}: {e}")
                continue
            image_inputs = clip_processor(images=image, return_tensors="pt")
            with torch.no_grad():
                image_features = clip_model.get_image_features(**image_inputs).to(device)
            image_features = image_features / image_features.norm(dim=-1, keepdim=True)
            score = torch.cosine_similarity(text_features, image_features).item()
            st.write(f"Similarity score for '{filename}': {score:.2f}")
            if score > best_score:
                best_score = score
                best_image = image_path

    st.write("Selected best image:", best_image, "with score:", best_score)
    return best_image

def create_final_presentation(slides, content_template_path, presentation_title, institution):
    """
    Create the final presentation using the content template (which defines the theme).
    The final presentation includes:
      1. A Title slide (built from the content template) updated with presentation_title and institution.
      2. One slide per GPT-generated content slide (with text on the left and best-fit image on the right).
      3. A Thank You slide.
    All slides use the theme defined in the content template.
    """
    # Load the content template; this preserves its theme and master.
    final_pres = Presentation(content_template_path)
    # Use the first slide as the Title slide.
    title_slide = final_pres.slides[0]
    # Clear any existing text (if necessary) and add title & institution.
    # Here, we assume the title slide will have two textboxes.
    # Remove all shapes from the title slide (except the background) if needed.
    # For simplicity, we add two new textboxes.
    slide_width_inches = final_pres.slide_width / 914400

    # Add Title textbox.
    title_box = title_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(slide_width_inches - 2), Inches(1.5))
    title_box.text_frame.text = presentation_title
    title_box.text_frame.word_wrap = True
    # Add Institution textbox.
    inst_box = title_slide.shapes.add_textbox(Inches(1), Inches(2.7), Inches(slide_width_inches - 2), Inches(1))
    inst_box.text_frame.text = institution
    inst_box.text_frame.word_wrap = True

    # --- Add Content Slides ---
    # Define layout: left 60% for text, right 40% for image.
    left_margin = 0.5
    right_margin = 0.5
    gap = 0.5
    available_width = slide_width_inches - left_margin - right_margin
    text_width = (available_width * 0.6) - (gap / 2)
    image_width = (available_width * 0.4) - (gap / 2)
    text_left = left_margin
    image_left = left_margin + text_width + gap

    # For each GPT slide, add a new slide using a blank layout from the content template.
    # We'll use slide_layouts[6] (a blank layout) if available; otherwise, use the first layout.
    try:
        blank_layout = final_pres.slide_layouts[6]
    except Exception:
        blank_layout = final_pres.slide_layouts[0]

    for slide_data in slides:
        slide = final_pres.slides.add_slide(blank_layout)
        # Add Title textbox.
        tbox = slide.shapes.add_textbox(Inches(text_left), Inches(0.5), Inches(text_width), Inches(1))
        tbox.text_frame.text = slide_data.get("title", "")
        tbox.text_frame.word_wrap = True
        # Add Content textbox.
        cbox = slide.shapes.add_textbox(Inches(text_left), Inches(1.8), Inches(text_width), Inches(3))
        cbox.text_frame.text = slide_data.get("content", "")
        cbox.text_frame.word_wrap = True
        # Insert best-fit image if available.
        best_image = select_best_image_for_slide(slide_data)
        if best_image and os.path.exists(best_image):
            slide.shapes.add_picture(best_image, Inches(image_left), Inches(1), width=Inches(image_width))

    # --- Add Thank You Slide ---
    thankyou_slide = final_pres.slides.add_slide(blank_layout)
    ty_box = thankyou_slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(slide_width_inches - 2), Inches(1.5))
    ty_box.text_frame.text = "Thank You"
    ty_box.text_frame.word_wrap = True

    return final_pres

def main():
    st.title("AI-Powered PowerPoint Presentation Generator")
    st.write(
        "Enter your presentation instructions below. The final presentation will include a Title slide (with your presentation title and university/college), "
        "content slides generated from your input (with text on the left and a best-fit image on the right), and a Thank You slide at the end. "
        "All slides will use the theme defined in your content template (powerpointtemplate.pptx)."
    )

    # Fields for Title slide information.
    presentation_title = st.text_input("Presentation Title", "My Presentation")
    institution = st.text_input("University/College", "ABC College")

    content_template_path = get_content_template()
    if content_template_path is None:
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

        with st.spinner("Creating final presentation..."):
            final_pres = create_final_presentation(slides, content_template_path, presentation_title, institution)
            if final_pres is None:
                st.error("Error creating final presentation.")
                return

            pptx_io = BytesIO()
            final_pres.save(pptx_io)
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
