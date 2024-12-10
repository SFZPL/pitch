import streamlit as st
import openai
import os
import logging
from datetime import datetime
import re
import time
from typing import List, Optional, Dict, Any
from PyPDF2 import PdfReader
from docx import Document
from pptx import Presentation
from pptx.util import Pt, Inches
import logging.handlers
import tiktoken
import base64
from io import BytesIO
from pptx.dml.color import RGBColor

# Logging Configuration
def setup_logging():
    log_formatter = logging.Formatter('%(asctime)s - %(levelname)s: %(message)s')
    log_file = f'pitch_deck_logs_{datetime.now().strftime("%Y%m%d")}.log'

    # Rotating File Handler
    file_handler = logging.handlers.RotatingFileHandler(
        log_file,
        maxBytes=10*1024*1024,  # 10MB
        backupCount=5
    )
    file_handler.setFormatter(log_formatter)

    # Console Handler
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(log_formatter)

    # Configure root logger
    logging.getLogger().addHandler(file_handler)
    logging.getLogger().addHandler(console_handler)
    logging.getLogger().setLevel(logging.INFO)

# Initialize logging
setup_logging()

# Secure OpenAI API Configuration using Streamlit Secrets
try:
    openai.api_key = st.secrets["OPENAI_API_KEY"]
    if not openai.api_key:
        raise ValueError("OpenAI API Key is missing")
    # Validate API key
    openai.Model.list()
except Exception as e:
    logging.error(f"OpenAI API Configuration Error: {e}")
    raise

class PitchDeckGenerator:
    def __init__(self):
        """Initialize the Pitch Deck Generator with system prompt and configuration."""
        self.system_prompt = """
You are an expert pitch deck content generator trained in creating world-class business presentations.
Your goal is to transform raw content into a compelling narrative that captures investor attention.

Key Presentation Principles:
1. Tell a Story: Create a clear, engaging narrative
2. Highlight Problem and Solution
3. Demonstrate Market Opportunity
4. Showcase Unique Value Proposition
5. Use Data-Driven Insights
6. Maintain Professional and Exciting Tone

Structure Recommendations:
- Title Slide: Company Name, Tagline
- Problem Statement: Clear, Impactful
- Solution: Innovative Approach
- Market Analysis: Size, Growth, Opportunity
- Business Model
- Competitive Landscape
- Financial Projections
- Team Overview
- Call to Action

Output Format:
For each slide, provide:
1. Slide Title
2. Content
3. Suggested Visual Representation
"""
        # Allowed file types
        self.ALLOWED_FILE_TYPES = [
            "text/plain",
            "application/pdf",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        ]

    def sanitize_input(self, text: str) -> str:
        return re.sub(r'[<>&\'"]', '', text)

    def clean_text(self, text: str) -> str:
        text = re.sub(r'-\n', '', text)
        text = re.sub(r'\n+', '\n', text)
        text = re.sub(r'[ \t]+', ' ', text)
        text = ''.join(char for char in text if char.isprintable())
        return text.strip()

    def count_tokens(self, text: str, model: str = "gpt-4") -> int:
        try:
            encoding = tiktoken.encoding_for_model(model)
            return len(encoding.encode(str(text)))
        except Exception as e:
            logging.error(f"Token counting error: {e}")
            return 0

    def extract_text_from_file(self, file):
        if file.type not in self.ALLOWED_FILE_TYPES:
            st.error(f"❌ Unsupported file type: {file.type}")
            return None

        try:
            content = ""
            file_type = file.type

            if file_type == "text/plain":
                content = file.read().decode("utf-8")
            elif file_type == "application/pdf":
                pdf_reader = PdfReader(file)
                content = "\n".join([page.extract_text() for page in pdf_reader.pages])
            elif file_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                doc = Document(file)
                content = "\n".join([para.text for para in doc.paragraphs if para.text])
            elif file_type == "application/vnd.openxmlformats-officedocument.presentationml.presentation":
                ppt = Presentation(file)
                content = "\n".join([
                    paragraph.text
                    for slide in ppt.slides
                    for shape in slide.shapes
                    if shape.has_text_frame
                    for paragraph in shape.text_frame.paragraphs
                ])
            else:
                st.warning("⚠️ Unsupported file type")
                return None

            return self.clean_text(content)

        except Exception as e:
            logging.error(f"File extraction error: {e}")
            st.error(f"❌ Error processing file: {e}")
            return None

    def summarize_content(self, content: str, max_tokens: int = 2048) -> Optional[str]:
        try:
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant that summarizes text."},
                    {"role": "user", "content": f"Please provide a concise summary of the following content:\n{content}"}
                ],
                max_tokens=max_tokens,
                temperature=0.5,
            )
            summary = response.choices[0].message.content
            return summary.strip() if summary else None
        except Exception as e:
            logging.error(f"Content summarization error: {e}")
            st.error(f"❌ Error summarizing content: {e}")
            return None

    def generate_pitch_deck(
        self,
        content: str,
        additional_context: str = "",
        previous_slides: Optional[List[str]] = None,
        max_retries: int = 3
    ) -> Optional[str]:
        for attempt in range(max_retries):
            try:
                context_prefix = ""
                if previous_slides:
                    context_prefix = "Previous Pitch Deck Context:\n" + "\n\n".join(previous_slides)

                content = self.sanitize_input(content)
                additional_context = self.sanitize_input(additional_context)

                full_context = f"{context_prefix}\n{additional_context}"

                total_input_tokens = (
                    self.count_tokens(self.system_prompt) +
                    self.count_tokens(full_context) +
                    self.count_tokens(content)
                )

                max_allowed_tokens = 7000

                if total_input_tokens > max_allowed_tokens:
                    st.warning("⚠️ Content is too long, summarizing to fit within token limits...")
                    content = self.summarize_content(content)
                    if not content:
                        return None

                    total_input_tokens = (
                        self.count_tokens(self.system_prompt) +
                        self.count_tokens(full_context) +
                        self.count_tokens(content)
                    )

                messages = [
                    {"role": "system", "content": self.system_prompt},
                    {"role": "user", "content": f"{full_context}\n\nDocument Content:\n{content}"}
                ]

                response = openai.ChatCompletion.create(
                    model="gpt-4",
                    messages=messages,
                    max_tokens=2500,
                    temperature=0.7
                )

                generated_content = response.choices[0].message.content
                logging.info("Pitch deck successfully generated")
                return generated_content

            except openai.error.RateLimitError:
                if attempt < max_retries - 1:
                    logging.warning(f"Rate limit reached. Retry attempt {attempt + 1} 💫")
                    time.sleep((attempt + 1) * 2)
                else:
                    logging.error("Max retries exceeded for pitch deck generation")
                    st.error("❌ Unable to generate pitch deck due to rate limits. Please try again later.")
                    return None
            except Exception as e:
                logging.error(f"Pitch deck generation error: {e}")
                st.error(f"❌ Error generating pitch deck: {e}")
                return None

    def update_slide(
        self,
        slide_number: int,
        slide_content: str,
        edit_instructions: str,
        previous_slides: List[str]
    ) -> Optional[str]:
        try:
            edit_instructions = self.sanitize_input(edit_instructions)
            slide_content = self.sanitize_input(slide_content)

            messages = [
                {"role": "system", "content": self.system_prompt},
                {"role": "user", "content": f"You are to update Slide {slide_number} of the pitch deck."}
            ]
            messages.append({"role": "user", "content": f"Previous Slides:\n" + "\n\n".join(previous_slides)})
            messages.append({"role": "user", "content": f"Current Slide {slide_number} Content:\n{slide_content}"})
            messages.append({"role": "user", "content": f"Edit Instructions:\n{edit_instructions}"})

            response = openai.ChatCompletion.create(
                model="gpt-4",
                messages=messages,
                max_tokens=500,
                temperature=0.7
            )
            updated_slide = response.choices[0].message.content
            logging.info(f"Slide {slide_number} updated successfully")
            return updated_slide.strip()
        except Exception as e:
            logging.error(f"Slide update error: {e}")
            st.error(f"❌ Error updating slide: {e}")
            return None

    def hex_to_rgb(self, hex_color):
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

    def export_to_pptx(self, slides: List[str]) -> Optional[BytesIO]:
        try:
            prs = Presentation()
            prs.slide_width = Inches(16)
            prs.slide_height = Inches(9)

            colors = {
                'background': 'F0F4F8',
                'primary': '1E40AF',
                'secondary': '3B82F6',
                'accent': 'EF4444',
            }

            for slide_content in slides:
                slide_layout = prs.slide_layouts[6]
                slide = prs.slides.add_slide(slide_layout)
                background = slide.background
                fill = background.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(*self.hex_to_rgb(colors['background']))

                parts = slide_content.split('\n', 2)

                if len(parts) > 0 and parts[0].strip():
                    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(15), Inches(1.5))
                    title_frame = title_box.text_frame
                    title_frame.text = parts[0]
                    title_frame.paragraphs[0].font.size = Pt(44)
                    title_frame.paragraphs[0].font.color.rgb = RGBColor(*self.hex_to_rgb(colors['primary']))
                    title_frame.paragraphs[0].font.bold = True

                if len(parts) > 1 and parts[1].strip():
                    content_box = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(15), Inches(6))
                    content_frame = content_box.text_frame
                    content_frame.text = parts[1]
                    content_frame.paragraphs[0].font.size = Pt(24)
                    content_frame.paragraphs[0].font.color.rgb = RGBColor(*self.hex_to_rgb(colors['secondary']))

            pptx_file = BytesIO()
            prs.save(pptx_file)
            pptx_file.seek(0)
            return pptx_file

        except Exception as e:
            logging.error(f"PowerPoint export error: {e}")
            st.error(f"❌ Failed to export PowerPoint: {e}")
            return None

def main():
    st.set_page_config(page_title="Prez.AI's Pitch Deck Generator", page_icon="📊", layout="wide")

    st.markdown("""
    <style>
    .main-header {
        font-size: 2.5em;
        color: #333333;
        text-align: center;
        margin-bottom: 20px;
        font-weight: bold;
    }
    .subheader {
        color: #666666;
        text-align: center;
        margin-bottom: 30px;
        font-size: 1.2em;
    }
    .custom-text {
        background-color: #f9f9f9;
        border-radius: 5px;
        padding: 15px;
        border: 1px solid #e0e0e0;
    }
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        font-weight: bold;
        border: none;
        padding: 10px 20px;
        font-size: 16px;
        margin: 4px 2px;
        transition-duration: 0.4s;
        cursor: pointer;
    }
    .stButton>button:hover {
        background-color: #45a049;
    }
    .chat-message {
        margin: 10px 0;
        padding: 10px;
        border-radius: 5px;
        font-size: 14px;
    }
    .user { background-color: #f0f0f0; }
    .assistant { background-color: #e6f3ff; }
    .system { background-color: #e6ffe6; }
    .error { background-color: #ffe6e6; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="main-header">Prez.AI Pitch Deck Generator 🎨📈</div>', unsafe_allow_html=True)
    st.markdown('<div class="subheader">Transform Your Ideas into Compelling Presentations 🚀</div>', unsafe_allow_html=True)

    if 'pitch_generator' not in st.session_state:
        st.session_state.pitch_generator = PitchDeckGenerator()

    session_defaults = {
        'uploaded_file': None,
        'pitch_deck_slides': [],
        'last_additional_context': "",
        'chat_history': [],
        'current_step': 'initial',
        'file_content': None
    }

    for key, default_value in session_defaults.items():
        if key not in st.session_state:
            st.session_state[key] = default_value

    col1, col2 = st.columns([1, 2])

    with col1:
        st.header("📄 Upload & Configure")

        uploaded_file = st.file_uploader(
            "📂 Upload Your Source Document",
            type=["txt", "pdf", "docx", "pptx"],
            help="Supported formats: Text, PDF, Word, PowerPoint"
        )

        if uploaded_file != st.session_state.uploaded_file:
            st.session_state.uploaded_file = uploaded_file
            st.session_state.pitch_deck_slides = []
            st.session_state.chat_history = []
            st.session_state.file_content = None
            st.session_state.current_step = 'file_uploaded'

        st.subheader("💡 Additional Context")
        additional_context = st.text_area(
            "Provide specific pitch deck guidelines:",
            value=st.session_state.last_additional_context,
            placeholder="e.g., Target audience, preferred tone, key messages...",
            height=150
        )
        st.session_state.last_additional_context = additional_context

        generate_button = st.button("✨ Generate Deck", type="primary", key="generate_button")
        clear_button = st.button("🧹 Clear Session", key="clear_button")

        if clear_button:
            for key in session_defaults:
                st.session_state[key] = session_defaults[key]
            st.experimental_rerun()

    with col2:
        st.header("🎬 Generated Pitch Deck")

        if not st.session_state.uploaded_file:
            st.info("📄 Upload a document to start generating your pitch deck.")
        elif st.session_state.uploaded_file and not st.session_state.pitch_deck_slides:
            if generate_button:
                with st.spinner("🔄 Analyzing document and generating pitch deck..."):
                    file_content = st.session_state.pitch_generator.extract_text_from_file(st.session_state.uploaded_file)
                    if file_content:
                        st.session_state.file_content = file_content
                        pitch_deck = st.session_state.pitch_generator.generate_pitch_deck(
                            content=file_content,
                            additional_context=st.session_state.last_additional_context
                        )
                        if pitch_deck:
                            st.session_state.pitch_deck_slides = pitch_deck.split("\n\n")
                            st.session_state.chat_history.append({
                                'type': 'system',
                                'message': "✅ Pitch Deck Generated Successfully!"
                            })
                            st.session_state.current_step = 'deck_generated'
                        else:
                            st.session_state.chat_history.append({
                                'type': 'error',
                                'message': "❌ Failed to generate pitch deck. Please check your file and try again."
                            })
                    else:
                        st.error("❌ Could not extract content from the file.")

        if st.session_state.pitch_deck_slides:
            slide_tabs = st.tabs([f"Slide {i+1}" for i in range(len(st.session_state.pitch_deck_slides))])

            for i, slide_content in enumerate(st.session_state.pitch_deck_slides):
                with slide_tabs[i]:
                    slide_parts = slide_content.split('\n', 2)
                    slide_title = slide_parts[0] if len(slide_parts) > 0 else ""
                    slide_body = slide_parts[1] if len(slide_parts) > 1 else ""
                    slide_visual = slide_parts[2] if len(slide_parts) > 2 else ""

                    st.subheader(slide_title)
                    st.markdown(slide_body, unsafe_allow_html=True)
                    st.markdown(f"{slide_visual}")

                    with st.expander("✏️ Edit This Slide"):
                        edit_instructions = st.text_area(
                            f"Modify Slide {i+1}",
                            placeholder="Adjust content, tone, or add specific details...",
                            key=f"edit_instructions_{i}"
                        )
                        if st.button(f"💾 Update Slide {i+1}", key=f"update_slide_{i}"):
                            if edit_instructions:
                                st.session_state.chat_history.append({
                                    'type': 'user',
                                    'message': f"✍️ Edit Slide {i+1}: {edit_instructions}"
                                })
                                with st.spinner("🔧 Updating slide..."):
                                    updated_slide = st.session_state.pitch_generator.update_slide(
                                        slide_number=i+1,
                                        slide_content=slide_content,
                                        edit_instructions=edit_instructions,
                                        previous_slides=st.session_state.pitch_deck_slides
                                    )
                                    if updated_slide:
                                        st.session_state.pitch_deck_slides[i] = updated_slide
                                        st.session_state.chat_history.append({
                                            'type': 'assistant',
                                            'message': f"✅ Updated Slide {i+1}:\n{updated_slide}"
                                        })
                                        st.success(f"Slide {i+1} updated successfully!")
                                        st.experimental_rerun()
                                    else:
                                        st.error("❌ Failed to update the slide.")
                            else:
                                st.warning("⚠️ Please provide edit instructions.")

            col2_1, col2_2 = st.columns(2)
            with col2_1:
                st.download_button(
                    label="💾 Download as Text",
                    data="\n\n".join(st.session_state.pitch_deck_slides),
                    file_name="pitch_deck.txt",
                    mime="text/plain"
                )
            with col2_2:
                export_pptx = st.button("📂 Export as PowerPoint")
                if export_pptx:
                    pptx_file = st.session_state.pitch_generator.export_to_pptx(st.session_state.pitch_deck_slides)
                    if pptx_file:
                        b64 = base64.b64encode(pptx_file.getvalue()).decode()
                        href = f'<a href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}" download="pitch_deck.pptx">📥 Download PowerPoint</a>'
                        st.markdown(href, unsafe_allow_html=True)

    st.markdown("---")
    st.subheader("📜 Session History")
    chat_container = st.container()

    with chat_container:
        for chat_entry in st.session_state.chat_history:
            message_class = chat_entry.get('type', 'user')
            st.markdown(f'<div class="chat-message {message_class}">{chat_entry["message"]}</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
