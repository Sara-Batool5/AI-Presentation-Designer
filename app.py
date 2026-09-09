import streamlit as st
from groq import Groq
import json
import os
import requests
import io
import pypdf
from PIL import Image

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

# Page Configuration
st.set_page_config(page_title="I Present - AI Presentation Designer", layout="wide", page_icon="🎨")

st.title("🎨 I Present - AI Presentation Designer")
st.subheader("Generate Visually Stunning PPTX Slides with Images & Custom Themes")

# Sidebar - API Key & Theme Configuration
st.sidebar.header("⚙️ Configuration & Styling")
api_key_input = st.sidebar.text_input("Enter your Groq API Key:", type="password")
api_key = api_key_input or os.environ.get("GROQ_API_KEY")

if not api_key:
    st.sidebar.warning("⚠️ Enter Groq API Key to enable generation.")

# Predefined Color Themes
THEMES = {
    "Corporate Navy": {
        "bg": RGBColor(248, 250, 252),        # Slate 50
        "primary": RGBColor(15, 23, 42),      # Slate 900
        "accent": RGBColor(14, 165, 233),     # Sky 500
        "card_bg": RGBColor(255, 255, 255),   # White
        "card_border": RGBColor(226, 232, 240),# Slate 200
        "text_dark": RGBColor(30, 41, 59),    # Slate 800
        "text_muted": RGBColor(100, 116, 139) # Slate 500
    },
    "Modern Emerald": {
        "bg": RGBColor(240, 253, 244),       # Mint 50
        "primary": RGBColor(6, 78, 59),        # Emerald 900
        "accent": RGBColor(16, 185, 129),     # Emerald 500
        "card_bg": RGBColor(255, 255, 255),
        "card_border": RGBColor(187, 247, 208),
        "text_dark": RGBColor(20, 83, 45),
        "text_muted": RGBColor(71, 85, 105)
    },
    "Midnight Cyber": {
        "bg": RGBColor(15, 23, 42),         # Dark Slate
        "primary": RGBColor(248, 250, 252),   # Off-white
        "accent": RGBColor(168, 85, 247),     # Purple 500
        "card_bg": RGBColor(30, 41, 59),      # Slate 800
        "card_border": RGBColor(51, 65, 85),  # Slate 700
        "text_dark": RGBColor(241, 245, 249),
        "text_muted": RGBColor(148, 163, 184)
    },
    "Warm Terracotta": {
        "bg": RGBColor(254, 252, 232),       # Amber 50
        "primary": RGBColor(120, 53, 15),     # Amber 900
        "accent": RGBColor(245, 158, 11),     # Amber 500
        "card_bg": RGBColor(255, 255, 255),
        "card_border": RGBColor(254, 215, 170),
        "text_dark": RGBColor(69, 26, 3),
        "text_muted": RGBColor(120, 53, 15)
    }
}

selected_theme_name = st.sidebar.selectbox("Choose Visual Theme:", list(THEMES.keys()))
enable_images = st.sidebar.checkbox("Include Relevant Stock Images", value=True)

# Helper: Extract text from PDF
def extract_text_from_pdf(pdf_file):
    reader = pypdf.PdfReader(pdf_file)
    text = ""
    for page in reader.pages:
        text += page.extract_text() or ""
    return text

# Helper: Fetch royalty-free image based on query
def fetch_stock_image(query):
    if not query:
        return None
    try:
        url = f"https://unsplash.com/napi/search/photos?query={query}&per_page=1"
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        res = requests.get(url, headers=headers, timeout=5)
        if res.status_code == 200:
            data = res.json()
            results = data.get("results", [])
            if results:
                img_url = results[0]["urls"]["regular"]
                img_res = requests.get(img_url, timeout=5)
                if img_res.status_code == 200:
                    return io.BytesIO(img_res.content)
    except Exception as e:
        print(f"Failed to fetch image for '{query}': {e}")
    return None

# Helper: Create PPTX with styled layouts and images
def create_designed_pptx(slides_data, presentation_title, theme_colors, include_images):
    prs = Presentation()
    prs.slide_width = Inches(13.333)  # 16:9 Widescreen ratio
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # --- 1. TITLE SLIDE ---
    title_slide = prs.slides.add_slide(blank_layout)
    
    # Background fill
    bg = title_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = theme_colors["bg"]
    bg.line.fill.background()

    # Title Card Container
    card = title_slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), Inches(1.8), Inches(10.333), Inches(3.8))
    card.fill.solid()
    card.fill.fore_color.rgb = theme_colors["card_bg"]
    card.line.color.rgb = theme_colors["accent"]
    card.line.width = Pt(2)

    # Title Text
    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.5)
    tf.margin_top = Inches(0.8)
    
    p = tf.paragraphs[0]
    p.text = presentation_title
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = theme_colors["primary"]
    p.alignment = PP_ALIGN.CENTER

    p2 = tf.add_paragraph()
    p2.text = "\nDesigned by I Present AI"
    p2.font.size = Pt(20)
    p2.font.color.rgb = theme_colors["accent"]
    p2.alignment = PP_ALIGN.CENTER

    # --- 2. CONTENT SLIDES ---
    for idx, slide_info in enumerate(slides_data):
        slide = prs.slides.add_slide(blank_layout)

        # Slide Background
        s_bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        s_bg.fill.solid()
        s_bg.fill.fore_color.rgb = theme_colors["bg"]
        s_bg.line.fill.background()

        # Top Accent Line
        accent_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.5), Inches(0.15), Inches(0.8))
        accent_bar.fill.solid()
        accent_bar.fill.fore_color.rgb = theme_colors["accent"]
        accent_bar.line.fill.background()

        # Slide Header
        tb = slide.shapes.add_textbox(Inches(1.1), Inches(0.4), Inches(11), Inches(0.9))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = slide_info.get("slide_title", f"Slide {idx + 1}")
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = theme_colors["primary"]

        # Fetch image if enabled
        img_query = slide_info.get("image_search_query", "")
        img_stream = fetch_stock_image(img_query) if include_images and img_query else None

        # Determine Layout based on Image Availability
        if img_stream:
            # Layout: Text Card (Left) + Image Card (Right)
            # Text Card Box
            text_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(7.2), Inches(5.3))
            text_card.fill.solid()
            text_card.fill.fore_color.rgb = theme_colors["card_bg"]
            text_card.line.color.rgb = theme_colors["card_border"]

            tf_text = text_card.text_frame
            tf_text.word_wrap = True
            tf_text.margin_left = Inches(0.4)
            tf_text.margin_top = Inches(0.4)

            for i, point in enumerate(slide_info.get("bullet_points", [])):
                p = tf_text.add_paragraph() if i > 0 else tf_text.paragraphs[0]
                p.text = f"• {point}"
                p.font.size = Pt(18)
                p.font.color.rgb = theme_colors["text_dark"]
                p.space_after = Pt(14)

            # Image Container
            try:
                slide.shapes.add_picture(img_stream, Inches(8.3), Inches(1.5), width=Inches(4.2), height=Inches(5.3))
            except Exception:
                pass  # Fallback gracefully if image parsing fails
        else:
            # Layout: Full Width Text Container
            full_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(11.733), Inches(5.3))
            full_card.fill.solid()
            full_card.fill.fore_color.rgb = theme_colors["card_bg"]
            full_card.line.color.rgb = theme_colors["card_border"]

            tf_text = full_card.text_frame
            tf_text.word_wrap = True
            tf_text.margin_left = Inches(0.5)
            tf_text.margin_top = Inches(0.5)

            for i, point in enumerate(slide_info.get("bullet_points", [])):
                p = tf_text.add_paragraph() if i > 0 else tf_text.paragraphs[0]
                p.text = f"• {point}"
                p.font.size = Pt(20)
                p.font.color.rgb = theme_colors["text_dark"]
                p.space_after = Pt(18)

    # Save to BytesIO
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# --- APP UI ---
st.header("1. Presentation Details & Input")

input_type = st.radio("Select Input Source:", ["Topic Name", "Upload PDF Document"], horizontal=True)

topic_or_context = ""
if input_type == "Topic Name":
    topic_or_context = st.text_input("Enter Presentation Topic:", placeholder="e.g., The Future of Artificial Intelligence in Healthcare")
else:
    uploaded_file = st.file_uploader("Upload Source PDF", type=["pdf"])
    if uploaded_file is not None:
        with st.spinner("Processing PDF content..."):
            topic_or_context = extract_text_from_pdf(uploaded_file)
            st.success("PDF processed successfully!")

col1, col2 = st.columns(2)
with col1:
    num_slides = st.number_input("Number of Content Slides:", min_value=3, max_value=15, value=5)
with col2:
    tone_option = st.selectbox("Presentation Tone:", ["Professional", "Formal", "Persuasive", "Educational", "Casual", "Custom"])

tone = st.text_input("Custom Tone:", value="Creative and Energetic") if tone_option == "Custom" else tone_option
extra_instructions = st.text_area("Extra Directives / Key Highlights:", placeholder="e.g., Include market statistics, keep bullet points punchy and concise.")

generate_button = st.button("🚀 Generate Presentation Deck", type="primary")

# Generation Logic
if generate_button:
    if not api_key:
        st.error("Please enter a Groq API Key in the sidebar.")
    elif not topic_or_context.strip():
        st.error("Please provide a topic or upload a PDF document.")
    else:
        with st.spinner("AI is crafting content, structure, and image queries..."):
            try:
                client = Groq(api_key=api_key)
                
                prompt = f"""
                You are a top-tier management consultant and presentation designer.
                Create a high-impact presentation based on:
                - Source Content: {topic_or_context[:6000]}
                - Slide Count: {num_slides}
                - Tone: {tone}
                - Extra Instructions: {extra_instructions}

                Output strictly valid JSON with this exact schema:
                {{
                    "presentation_title": "Catchy Title",
                    "slides": [
                        {{
                            "slide_number": 1,
                            "slide_title": "Slide Title",
                            "bullet_points": ["Key point 1", "Key point 2", "Key point 3"],
                            "image_search_query": "2-3 word query for finding relevant image e.g. 'hospital robot doctor'"
                        }}
                    ],
                    "short_speaker_notes": ["Slide 1 notes...", "Slide 2 notes..."],
                    "q_and_a": [
                        {{"question": "Audience Question?", "answer": "Answer..."}}
                    ]
                }}
                Do not wrap response in markdown blocks. Output pure JSON only.
                """

                response = client.chat.completions.create(
                    messages=[{"role": "user", "content": prompt}],
                    model="openai/gpt-oss-120b",
                    temperature=0.7
                )
                
                raw_txt = response.choices[0].message.content.strip()
                
                # Strip markdown code blocks if present
                if raw_txt.startswith("```json"):
                    raw_txt = raw_txt[7:]
                elif raw_txt.startswith("```"):
                    raw_txt = raw_txt[3:]
                    
                if raw_txt.endswith("```"):
                    raw_txt = raw_txt[:-3]
                
                raw_txt = raw_txt.strip()
                data = json.loads(raw_txt)
                
                data = json.loads(raw_txt.strip())
                st.session_state['generated_data'] = data
                
                # Create designed PPTX file
                theme_colors = THEMES[selected_theme_name]
                pptx_output = create_designed_pptx(data.get("slides", []), data.get("presentation_title", "Presentation"), theme_colors, enable_images)
                st.session_state['pptx_file'] = pptx_output
                st.success("✨ Presentation designed and generated successfully!")
                
            except Exception as e:
                st.error(f"Error during presentation generation: {str(e)}")

# Display Outputs
if 'generated_data' in st.session_state:
    data = st.session_state['generated_data']
    pptx_file = st.session_state['pptx_file']
    
    st.divider()
    st.header("2. Your Generated Presentation Package")
    
    st.download_button(
        label="📥 Download Presentation (.pptx)",
        data=pptx_file,
        file_name=f"{data.get('presentation_title', 'presentation').lower().replace(' ', '_')}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
    
    tab1, tab2, tab3 = st.tabs(["🖼️ Slide Structure & Image Queries", "📝 Speaker Preparation Notes", "❓ Audience Q&A"])
    
    with tab1:
        st.subheader(data.get("presentation_title", "Presentation"))
        for s in data.get("slides", []):
            with st.expander(f"Slide {s.get('slide_number')}: {s.get('slide_title')}"):
                for bp in s.get("bullet_points", []):
                    st.write(f"• {bp}")
                if s.get("image_search_query"):
                    st.caption(f"🔍 Image Search Query used: *'{s.get('image_search_query')}'*")

    with tab2:
        st.subheader("Speaker Notes")
        for note in data.get("short_speaker_notes", []):
            st.info(note)

    with tab3:
        st.subheader("Anticipated Questions & Answers")
        for idx, qa in enumerate(data.get("q_and_a", []), 1):
            st.markdown(f"**Q{idx}: {qa.get('question')}**")
            st.write(f"**A:** {qa.get('answer')}")
            st.write("---")
