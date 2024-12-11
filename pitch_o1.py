import streamlit as st
import openai
import os
import logging
from datetime import datetime
import re
import time
from typing import List, Optional
from io import BytesIO
import base64
import logging.handlers

from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
import tiktoken

# --------------------------------------------
# Constants & Globals
# --------------------------------------------
ALLOWED_FILE_TYPES = [
    "text/plain",
    "application/pdf",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
]

# Model specifics for gpt-3.5-turbo-16k
MAX_CONTEXT = 16385  # model's max context length

def clean_text(text: str) -> str:
    text = re.sub(r'-\n', '', text)
    text = re.sub(r'\n+', '\n', text)
    text = re.sub(r'[ \t]+', ' ', text)
    text = ''.join(char for char in text if char.isprintable())
    return text.strip()

@st.cache_data
def count_tokens(text: str, model: str = "gpt-3.5-turbo-16k") -> int:
    try:
        encoding = tiktoken.encoding_for_model(model)
        return len(encoding.encode(str(text)))
    except Exception as e:
        logging.error(f"Token counting error: {e}")
        return 0

@st.cache_data
def extract_text_from_file(file) -> Optional[str]:
    if file.type not in ALLOWED_FILE_TYPES:
        st.error(f"❌ Unsupported file type: {file.type}")
        return None

    try:
        content = ""
        file_type = file.type

        if file_type == "text/plain":
            content = file.read().decode("utf-8")

        elif file_type == "application/pdf":
            from PyPDF2 import PdfReader
            pdf_reader = PdfReader(file)
            content = "\n".join([page.extract_text() for page in pdf_reader.pages if page.extract_text()])

        elif file_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            from docx import Document
            doc = Document(file)
            content = "\n".join([para.text for para in doc.paragraphs if para.text])

        elif file_type == "application/vnd.openxmlformats-officedocument.presentationml.presentation":
            from pptx import Presentation
            ppt = Presentation(file)
            content = "\n".join([
                paragraph.text
                for slide in ppt.slides
                for shape in slide.shapes
                if shape.has_text_frame
                for paragraph in shape.text_frame.paragraphs
            ])

        return clean_text(content)

    except Exception as e:
        logging.error(f"File extraction error: {e}")
        st.error(f"❌ Error processing file: {e}")
        return None

@st.cache_data
def summarize_content(content: str, max_tokens: int = 2048) -> Optional[str]:
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

def setup_logging():
    try:
        log_formatter = logging.Formatter('%(asctime)s - %(levelname)s: %(message)s')
        log_file = f'pitch_deck_logs_{datetime.now().strftime("%Y%m%d")}.log'

        # Rotating File Handler
        file_handler = logging.handlers.RotatingFileHandler(
            log_file,
            maxBytes=10 * 1024 * 1024,  # 10MB
            backupCount=5
        )
        file_handler.setFormatter(log_formatter)

        # Console Handler
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(log_formatter)

        # Configure root logger
        logger = logging.getLogger()
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        logger.setLevel(logging.INFO)

        logging.info("Logging successfully configured.")
    except Exception as e:
        print(f"Error setting up logging: {e}")
        raise

setup_logging()

try:
    openai.api_key = st.secrets["OPENAI_API_KEY"]
    openai.Model.list()
    logging.info("OpenAI API key validated successfully.")
except Exception as e:
    logging.error(f"OpenAI initialization error: {e}")
    st.error("❌ OpenAI API initialization error.")
    raise

class DocumentGenerator:
    def __init__(self):
        self.pitch_deck_prompt = """
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

        self.corporate_profile_prompt = """
You are a professional corporate profile content generator. Your task is to create a comprehensive and engaging corporate profile presentation (PowerPoint style) that effectively represents the company's brand and offering.

Key Elements to Include:
1. Title Slide: Company Name, Logo, and Tagline
2. Mission, Vision, and Values
3. Company History: Founding story, key milestones
4. Products/Services: Detailed descriptions, unique selling points
5. Market Position: Industry standing, competitive advantages
6. Achievements & Awards: Recognitions, milestones
7. Team & Leadership Overview
8. Future Goals & Strategic Plans
9. Contact Information

Output Format:
For each slide, provide:
1. Slide Title
2. Content (brief bullet points)
3. Suggested Visual Representation

Focus on professionalism, clarity, and cohesive design suitable for a corporate profile PowerPoint.
"""

    def sanitize_input(self, text: str) -> str:
        return re.sub(r'[<>&\'"]', '', text)

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

    def generate_document(
        self,
        content: str,
        additional_context: str = "",
        document_type: str = "Pitch Deck",
        max_retries: int = 3
    ) -> Optional[str]:
        # Choose the appropriate system prompt
        if document_type == "Pitch Deck":
            system_prompt = self.pitch_deck_prompt
        elif document_type == "Corporate Profile":
            system_prompt = self.corporate_profile_prompt
        else:
            st.error("❌ Unsupported document type selected.")
            return None

        # Initial desired completion tokens
        desired_completion_tokens = 8000

        for attempt in range(max_retries):
            try:
                content = self.sanitize_input(content)
                additional_context = self.sanitize_input(additional_context)
                full_context = f"{additional_context}"

                # Count tokens for the prompt
                prompt_tokens = (
                    count_tokens(system_prompt, model="gpt-3.5-turbo-16k") +
                    count_tokens(full_context, model="gpt-3.5-turbo-16k") +
                    count_tokens(content, model="gpt-3.5-turbo-16k")
                )

                # Check if it fits in max context
                if prompt_tokens + desired_completion_tokens > MAX_CONTEXT:
                    st.warning("⚠️ Content plus desired output too large, summarizing content...")
                    summarized = summarize_content(content)
                    if summarized:
                        content = summarized
                        # Recount after summarization
                        prompt_tokens = (
                            count_tokens(system_prompt, model="gpt-3.5-turbo-16k") +
                            count_tokens(full_context, model="gpt-3.5-turbo-16k") +
                            count_tokens(content, model="gpt-3.5-turbo-16k")
                        )
                        # If still too large, reduce desired_completion_tokens
                        if prompt_tokens + desired_completion_tokens > MAX_CONTEXT:
                            # Calculate a safe maximum completion tokens that fit in the context
                            safe_max = MAX_CONTEXT - prompt_tokens - 500  # a 500 token safety margin
                            if safe_max < 1000:
                                # If still too large, maybe warn user or force another summarization
                                st.warning("⚠️ Even after summarization, content is too large. Further reducing completion tokens.")
                                safe_max = max(500, safe_max) # ensure some reasonable output
                            desired_completion_tokens = safe_max
                    else:
                        # Summarization failed, cannot proceed
                        return None

                response = openai.ChatCompletion.create(
                    model="gpt-3.5-turbo-16k",
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": f"{full_context}\n\nDocument Content:\n{content}"}
                    ],
                    max_tokens=desired_completion_tokens,
                    temperature=0.7
                )

                generated_content = response.choices[0].message.content
                logging.info(f"{document_type} successfully generated.")
                return generated_content

            except openai.error.RateLimitError:
                if attempt < max_retries - 1:
                    logging.warning(f"Rate limit reached. Retry attempt {attempt + 1} 💫")
                    time.sleep((attempt + 1) * 2)
                else:
                    logging.error("Max retries exceeded for document generation.")
                    st.error("❌ Unable to generate document due to rate limits. Please try again later.")
                    return None
            except openai.error.InvalidRequestError as ire:
                logging.error(f"Invalid request error: {ire}")
                st.error(f"❌ Error generating {document_type}: {ire}")
                return None
            except Exception as e:
                logging.error(f"{document_type} generation error: {e}")
                st.error(f"❌ Error generating {document_type}: {e}")
                return None

    def update_section(
        self,
        section_number: int,
        section_content: str,
        edit_instructions: str,
        previous_sections: List[str],
        document_type: str
    ) -> Optional[str]:
        edit_instructions = self.sanitize_input(edit_instructions)
        section_content = self.sanitize_input(section_content)

        if document_type == "Pitch Deck":
            system_prompt = self.pitch_deck_prompt
        elif document_type == "Corporate Profile":
            system_prompt = self.corporate_profile_prompt
        else:
            st.error("❌ Unsupported document type selected.")
            return None

        try:
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo-16k",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": f"You are to update Section {section_number} of the {document_type}."},
                    {"role": "user", "content": f"Previous Sections:\n" + "\n\n".join(previous_sections)},
                    {"role": "user", "content": f"Current Section {section_number} Content:\n{section_content}"},
                    {"role": "user", "content": f"Edit Instructions:\n{edit_instructions}"}
                ],
                max_tokens=2000,
                temperature=0.7
            )
            updated_section = response.choices[0].message.content
            logging.info(f"Section {section_number} updated successfully.")
            return updated_section.strip()
        except Exception as e:
            logging.error(f"Section update error: {e}")
            st.error(f"❌ Error updating section: {e}")
            return None

    def analyze_existing_presentation(self, content: str, document_type: str) -> Optional[str]:
        # Provide an analysis of strengths, weaknesses, and suggestions
        analysis_prompt = f"""
You are a presentation analyst. The user has provided an existing {document_type}. 
Please analyze the content and provide:
- Key strengths: Which aspects are well done?
- Potential weaknesses or areas for improvement: What could be improved?
- Suggested tweaks or enhancements: How to strengthen the narrative, visuals, or clarity?

Keep the tone constructive and professional.
"""
        try:
            content = self.sanitize_input(content)
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo-16k",
                messages=[
                    {"role": "system", "content": analysis_prompt},
                    {"role": "user", "content": f"Existing {document_type} Content:\n{content}"}
                ],
                max_tokens=2000,
                temperature=0.7
            )
            analysis = response.choices[0].message.content
            return analysis.strip()
        except Exception as e:
            logging.error(f"Presentation analysis error: {e}")
            st.error(f"❌ Error analyzing presentation: {e}")
            return None

def main():
    st.set_page_config(page_title="Prez.AI Document Generator", page_icon="📊", layout="wide")

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

    st.markdown('<div class="main-header">Prez.AI Document Generator 🎨📈</div>', unsafe_allow_html=True)
    st.markdown('<div class="subheader">Transform Your Ideas into Compelling Documents 🚀</div>', unsafe_allow_html=True)

    if 'doc_generator' not in st.session_state:
        st.session_state.doc_generator = DocumentGenerator()

    session_defaults = {
        'uploaded_file': None,
        'generated_document_sections': [],
        'last_additional_context': "",
        'chat_history': [],
        'current_step': 'initial',
        'file_content': None,
        'document_type': "Pitch Deck",
        'analysis_result': None
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
            st.session_state.generated_document_sections = []
            st.session_state.chat_history = []
            st.session_state.file_content = None
            st.session_state.analysis_result = None
            st.session_state.current_step = 'file_uploaded'

        st.subheader("📝 Select Document Type")
        document_type = st.selectbox(
            "Choose the type of document you want to generate:",
            ["Pitch Deck", "Corporate Profile"],
            index=["Pitch Deck", "Corporate Profile"].index(st.session_state.document_type)
        )
        st.session_state.document_type = document_type

        st.subheader("💡 Additional Context")
        additional_context = st.text_area(
            "Provide specific guidelines or preferences:",
            value=st.session_state.last_additional_context,
            placeholder="e.g., Target audience, preferred tone, key messages...",
            height=150
        )
        st.session_state.last_additional_context = additional_context

        generate_button = st.button("✨ Generate Document", type="primary", key="generate_button")
        analyze_button = st.button("🔍 Analyze Existing Presentation", key="analyze_button")
        clear_button = st.button("🧹 Clear Session", key="clear_button")

        if clear_button:
            for key in session_defaults:
                st.session_state[key] = session_defaults[key]
            st.rerun()

    with col2:
        st.header("🎬 Generated Document")

        if st.session_state.uploaded_file and st.session_state.file_content is None:
            # Extract content if not already extracted
            st.session_state.file_content = extract_text_from_file(st.session_state.uploaded_file)

        if not st.session_state.uploaded_file:
            st.info("📄 Upload a document to start generating your presentation.")
        elif st.session_state.uploaded_file and not st.session_state.generated_document_sections and not st.session_state.analysis_result:
            # Show buttons after upload
            if generate_button:
                if st.session_state.file_content:
                    st.write("Now extracting text from file...")
                    file_content = st.session_state.file_content
                    if file_content:
                        st.write("The AI is cooking... please wait!")
                        doc = st.session_state.doc_generator.generate_document(
                            content=file_content,
                            additional_context=st.session_state.last_additional_context,
                            document_type=st.session_state.document_type
                        )
                        if doc:
                            st.write("Generation done!")
                            st.session_state.generated_document_sections = doc.split("\n\n")
                            st.session_state.chat_history.append({
                                'type': 'system',
                                'message': f"✅ {st.session_state.document_type} Generated Successfully!"
                            })
                            st.session_state.current_step = 'document_generated'
                        else:
                            st.session_state.chat_history.append({
                                'type': 'error',
                                'message': f"❌ Failed to generate {st.session_state.document_type}."
                            })
                    else:
                        st.error("❌ Could not extract content from the file.")

            if analyze_button:
                if st.session_state.file_content:
                    with st.spinner("Analyzing existing presentation..."):
                        analysis = st.session_state.doc_generator.analyze_existing_presentation(
                            content=st.session_state.file_content,
                            document_type=st.session_state.document_type
                        )
                        if analysis:
                            st.session_state.analysis_result = analysis
                            st.session_state.chat_history.append({
                                'type': 'assistant',
                                'message': "✅ Analysis Complete"
                            })
                        else:
                            st.error("❌ Failed to analyze the presentation.")


        if st.session_state.generated_document_sections:
            section_tabs = st.tabs([f"Section {i+1}" for i in range(len(st.session_state.generated_document_sections))])

            for i, section_content in enumerate(st.session_state.generated_document_sections):
                with section_tabs[i]:
                    parts = section_content.split('\n', 2)
                    slide_title = parts[0].strip() if len(parts) > 0 else "No Title"
                    slide_body = parts[1].strip() if len(parts) > 1 else ""
                    slide_visual = parts[2].strip() if len(parts) > 2 else ""

                    st.subheader(slide_title)
                    if slide_body:
                        st.markdown(slide_body, unsafe_allow_html=True)
                    if slide_visual:
                        st.markdown(slide_visual, unsafe_allow_html=True)

                    with st.expander("✏️ Edit This Section"):
                        edit_instructions = st.text_area(
                            f"Modify Section {i+1}",
                            placeholder="Adjust content, tone, or add specific details...",
                            key=f"edit_instructions_{i}"
                        )
                        if st.button(f"💾 Update Section {i+1}", key=f"update_section_{i}"):
                            if edit_instructions:
                                st.session_state.chat_history.append({
                                    'type': 'user',
                                    'message': f"✍️ Edit Section {i+1}: {edit_instructions}"
                                })
                                with st.spinner("🔧 Updating section..."):
                                    updated_section = st.session_state.doc_generator.update_section(
                                        section_number=i+1,
                                        section_content=section_content,
                                        edit_instructions=edit_instructions,
                                        previous_sections=st.session_state.generated_document_sections,
                                        document_type=st.session_state.document_type
                                    )
                                    if updated_section:
                                        st.session_state.generated_document_sections[i] = updated_section
                                        st.session_state.chat_history.append({
                                            'type': 'assistant',
                                            'message': f"✅ Updated Section {i+1}:\n{updated_section}"
                                        })
                                        st.success(f"Section {i+1} updated successfully!")
                                        st.rerun()
                                    else:
                                        st.error("❌ Failed to update the section.")
                            else:
                                st.warning("⚠️ Please provide edit instructions.")

            col2_1, col2_2 = st.columns(2)
            with col2_1:
                st.download_button(
                    label="💾 Download as Text",
                    data="\n\n".join(st.session_state.generated_document_sections),
                    file_name=f"{st.session_state.document_type.lower().replace(' ', '_')}.txt",
                    mime="text/plain"
                )
            with col2_2:
                export_pptx = st.button("📂 Export as PowerPoint")
                if export_pptx:
                    pptx_file = st.session_state.doc_generator.export_to_pptx(
                        slides=st.session_state.generated_document_sections
                    )
                    if pptx_file:
                        b64 = base64.b64encode(pptx_file.getvalue()).decode()
                        href = f'<a href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}" download="{st.session_state.document_type.lower().replace(" ", "_")}.pptx">📥 Download PowerPoint</a>'
                        st.markdown(href, unsafe_allow_html=True)

        if st.session_state.analysis_result:
            st.header("🔍 Presentation Analysis")
            st.markdown(st.session_state.analysis_result, unsafe_allow_html=True)

    st.markdown("---")
    st.subheader("📜 Session History")
    chat_container = st.container()

    with chat_container:
        for chat_entry in st.session_state.chat_history:
            message_class = chat_entry.get('type', 'user')
            st.markdown(f'<div class="chat-message {message_class}">{chat_entry["message"]}</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
