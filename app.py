import streamlit as st
from google import genai
from google.genai import types
from pypdf import PdfReader
from docx import Document
from pptx import Presentation
from gtts import gTTS
import graphviz
import io
import os
import json
import time

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="StudySphere", page_icon="logo.png", layout="wide")

# --- SECURE API KEY RETRIEVAL (SERVER-SIDE ONLY) ---
def get_server_api_key():
    """Retrieve Gemini API key from environment or secrets without exposing inputs in UI."""
    if "GEMINI_API_KEY" in os.environ:
        return os.environ["GEMINI_API_KEY"]
    try:
        if "GEMINI_API_KEY" in st.secrets:
            return st.secrets["GEMINI_API_KEY"]
    except Exception:
        pass
    return ""

API_KEY = get_server_api_key()

# --- VIVARILY NEON GLOW THEME & CRISP IMAGE RENDER CSS ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=Inter:wght@400;500;600&display=swap');

    .stApp {
        background-color: #0a050d !important;
        background-image: 
            radial-gradient(circle at 85% 30%, rgba(147, 51, 234, 0.25) 0%, rgba(236, 72, 153, 0.15) 35%, transparent 70%),
            linear-gradient(180deg, #0e0714 0%, #0a050d 100%) !important;
        background-attachment: fixed !important;
        color: #f3e8ff !important;
        font-family: 'Inter', sans-serif !important;
    }

    /* Force HD crisp rendering on all Streamlit images */
    img {
        image-rendering: -webkit-optimize-contrast !important;
        image-rendering: crisp-edges !important;
        object-fit: contain !important;
    }

    /* Circular Mask & Centered Alignment for Sidebar Logo Header */
    div[data-testid="stSidebarHeader"] {
        display: none !important;
    }

    .circle-logo-container {
        display: flex !important;
        justify-content: center !important;
        align-items: center !important;
        width: 100% !important;
        margin: 10px auto 15px auto !important;
    }

    .circle-logo-container img {
        border-radius: 50% !important;
        object-fit: cover !important;
        border: 2px solid rgba(217, 70, 239, 0.6) !important;
        box-shadow: 0px 0px 20px rgba(217, 70, 239, 0.5) !important;
    }

    h1 {
        font-family: 'Space Mono', monospace !important;
        color: #ffffff !important;
        font-size: 2.2rem !important;
        font-weight: 700 !important;
        letter-spacing: 2px !important;
        text-shadow: 0 0 12px rgba(217, 70, 239, 0.3) !important;
    }

    /* Refined Violet Headings */
    h2, h3, h4 {
        font-family: 'Space Mono', monospace !important;
        color: #c084fc !important;
        font-weight: 700 !important;
    }

    section[data-testid="stSidebar"] {
        background-color: #0f0817 !important;
        border-right: 1px solid #2e1065 !important;
    }

    p, span, label, div, caption, .stCaption, li {
        color: #e9d5ff !important;
    }

    .stButton>button {
        background: rgba(15, 8, 23, 0.7) !important;
        color: #c084fc !important;
        font-family: 'Space Mono', monospace !important;
        font-size: 0.85rem !important;
        font-weight: 700 !important;
        text-transform: uppercase !important;
        border: 1px solid #d946ef !important;
        border-radius: 4px !important;
        box-shadow: 0px 0px 10px rgba(217, 70, 239, 0.2) !important;
        transition: all 0.2s ease-in-out !important;
    }

    .stButton>button:hover {
        background: #d946ef !important;
        color: #0a050d !important;
        box-shadow: 0px 0px 18px rgba(217, 70, 239, 0.6) !important;
    }

    div[data-baseweb="select"] > div {
        background-color: #140b21 !important;
        border: 1px solid #a855f7 !important;
        color: #ffffff !important;
        font-family: 'Space Mono', monospace !important;
    }

    div[data-testid="stFileUploaderDropzone"] {
        background-color: #140b21 !important;
        border: 1px dashed #d946ef !important;
    }

    div[data-element-container] > div[data-testid="stVerticalBlock"] > div[style*="border"] {
        background: rgba(20, 11, 33, 0.6) !important;
        border: 1px solid #3b0764 !important;
        border-radius: 8px !important;
        backdrop-filter: blur(8px) !important;
    }

    div[data-testid="stExpander"] {
        background-color: #140b21 !important;
        border: 1px solid #3b0764 !important;
    }

    input, textarea {
        background-color: #0f0817 !important;
        border: 1px solid #581c87 !important;
        color: #ffffff !important;
        font-family: 'Space Mono', monospace !important;
    }

    div[data-testid="stChatInput"] textarea {
        background-color: #140b21 !important;
        border: 1px solid #d946ef !important;
        color: #ffffff !important;
    }

    div[data-testid="stChatMessage"] {
        background-color: #140b21 !important;
        border: 1px solid #3b0764 !important;
        border-radius: 6px !important;
    }

    div[data-testid="stChatMessage"]:nth-child(even) {
        background-color: #1e0a2f !important;
        border-color: #7e22ce !important;
    }

    header[data-testid="stHeader"] {
        background: transparent !important;
    }

    /* CYBER SEGMENTED CONTROL / TIMER & RIGOR PILLS */
    div[data-testid="stSegmentedControl"] {
        background-color: #0f0817 !important;
        border: 1px solid #3b0764 !important;
        border-radius: 8px !important;
        padding: 4px !important;
    }

    div[data-testid="stSegmentedControl"] button {
        background: transparent !important;
        color: #c084fc !important;
        font-family: 'Space Mono', monospace !important;
        font-size: 0.8rem !important;
        font-weight: 700 !important;
        border: none !important;
        border-radius: 6px !important;
        transition: all 0.2s ease-in-out !important;
    }

    div[data-testid="stSegmentedControl"] button[aria-selected="true"] {
        background: #d946ef !important;
        color: #0a050d !important;
        box-shadow: 0px 0px 15px rgba(217, 70, 239, 0.6) !important;
    }

    div[data-testid="stSegmentedControl"] button:hover:not([aria-selected="true"]) {
        background: rgba(217, 70, 239, 0.15) !important;
        color: #f472b6 !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- FILE PERSISTENCE & LOCAL HELPERS ---
SETTINGS_FILE = "settings.txt"
PHOTO_FILE = "profile_pic.png"

def load_local_settings():
    settings = {"user": "Active Student", "questions_answered": 0, "study_time_min": 0}
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                for line in f:
                    if "=" in line:
                        k, v = line.strip().split("=", 1)
                        if k in ["questions_answered", "study_time_min"]:
                            settings[k] = int(v)
                        else:
                            settings[k] = v
        except Exception:
            pass
    return settings

def save_local_settings(username, q_count, s_time):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            f.write(f"user={username}\n")
            f.write(f"questions_answered={q_count}\n")
            f.write(f"study_time_min={s_time}\n")
    except Exception:
        pass

def parse_uploaded_document(uploaded_file):
    """Extract text from PDF, DOCX, PPTX, or TXT notes."""
    text = ""
    filename = uploaded_file.name.lower()
    try:
        if filename.endswith(".pdf"):
            reader = PdfReader(io.BytesIO(uploaded_file.read()))
            for page in reader.pages:
                t = page.extract_text()
                if t:
                    text += t + "\n"
        elif filename.endswith(".docx"):
            doc = Document(io.BytesIO(uploaded_file.read()))
            for p in doc.paragraphs:
                text += p.text + "\n"
        elif filename.endswith(".pptx"):
            prs = Presentation(io.BytesIO(uploaded_file.read()))
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        text += shape.text + "\n"
        elif filename.endswith(".txt"):
            text = uploaded_file.read().decode("utf-8")
    except Exception as e:
        st.error(f"Error reading file {filename}: {e}")
    return text

# --- INITIALIZE STATE ---
saved_config = load_local_settings()

if "username" not in st.session_state:
    st.session_state.username = saved_config["user"]
if "questions_answered" not in st.session_state:
    st.session_state.questions_answered = saved_config["questions_answered"]
if "study_time_min" not in st.session_state:
    st.session_state.study_time_min = saved_config["study_time_min"]
if "uploaded_photo" not in st.session_state:
    if os.path.exists(PHOTO_FILE):
        try:
            with open(PHOTO_FILE, "rb") as f:
                st.session_state.uploaded_photo = f.read()
        except Exception:
            st.session_state.uploaded_photo = None
    else:
        st.session_state.uploaded_photo = None

if "chat_history" not in st.session_state:
    st.session_state.chat_history = []
if "assessment_questions" not in st.session_state:
    st.session_state.assessment_questions = []
if "assessment_evaluation" not in st.session_state:
    st.session_state.assessment_evaluation = None
if "mindmap_code" not in st.session_state:
    st.session_state.mindmap_code = None
if "flashcards_csv" not in st.session_state:
    st.session_state.flashcards_csv = None

# --- SIDEBAR CONFIGURATION ---
with st.sidebar:
    # Centered Large Circular Logo Header
    if os.path.exists("logo.png"):
        _, center_col, _ = st.columns([1, 4, 1])
        with center_col:
            st.markdown('<div class="circle-logo-container">', unsafe_allow_html=True)
            st.image("logo.png", use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
        
    st.caption("// EXPAND YOUR MIND BEYOND BOUNDARIES")
    st.divider()
    
    # Profile Card Container with Crisp Avatar Rendering
    with st.container(border=True):
        st.markdown("### 🧑‍🎓 Student Profile")
        col1, col2 = st.columns([1, 2])
        with col1:
            if st.session_state.uploaded_photo is not None:
                st.image(st.session_state.uploaded_photo, width=64)
            else:
                st.markdown("## ⚡")
        with col2:
            st.markdown(f"**{st.session_state.username}**")

    # Profile Settings Area
    with st.expander("⚙️ Profile Configuration"):
        new_name = st.text_input("Username", value=st.session_state.username)
        photo_file = st.file_uploader("Upload Profile Picture", type=["png", "jpg", "jpeg"])
        if st.button("Save Profile Settings", use_container_width=True):
            st.session_state.username = new_name
            if photo_file is not None:
                photo_bytes = photo_file.read()
                st.session_state.uploaded_photo = photo_bytes
                try:
                    with open(PHOTO_FILE, "wb") as f:
                        f.write(photo_bytes)
                except Exception:
                    pass
            save_local_settings(new_name, st.session_state.questions_answered, st.session_state.study_time_min)
            st.success("Profile Updated!")
            st.rerun()

    st.divider()
    
    # Learning Modes
    st.markdown("### 🎯 Learning Engine")
    study_mode = st.selectbox(
        "Select Active Mode:",
        [
            "📋 Practice Assessment & Evaluation",
            "📖 Lecture Mode (Deep Explanations)", 
            "⚡ Flashcard Maker (Anki Export)",
            "🧠 Mind Map Node Visualizer",
            "👁️ Vision Notes & Formula Reader"
        ]
    )
    
    # Academic Rigor Level (Segmented Control)
    st.markdown("### 🎓 Academic Rigor Level")
    academic_level = st.segmented_control(
        "Academic Rigor Level",
        options=["High School", "Undergraduate", "Postgraduate / PhD"],
        default="Undergraduate",
        key="rigor_control",
        label_visibility="collapsed"
    )
    if not academic_level:
        academic_level = "Undergraduate"
    
    st.divider()
    
    # Context Multi-Format File Upload
    st.markdown("### 📄 Multi-Format Context Material")
    uploaded_files = st.file_uploader(
        "Attach Course Notes (PDF, DOCX, PPTX, TXT)", 
        type=["pdf", "docx", "pptx", "txt"], 
        accept_multiple_files=True,
        key="doc_uploader"
    )
    
    document_text = ""
    if uploaded_files:
        for f in uploaded_files:
            document_text += f"\n--- File: {f.name} ---\n" + parse_uploaded_document(f)
        if document_text:
            st.success(f"Ingested {len(uploaded_files)} Document(s)!")

    st.divider()
    
    # Session Analytics Panel
    st.markdown("### 📊 Cyber Analytics")
    st.caption(f"Questions Evaluated: **{st.session_state.questions_answered}**")
    st.caption(f"Est. Study Time Saved: **{st.session_state.questions_answered * 4} mins**")

    st.divider()
    
    # Workspace Actions & Logs
    if st.button("🧹 Clear Workspace", use_container_width=True):
        st.session_state.chat_history = []
        st.session_state.assessment_questions = []
        st.session_state.assessment_evaluation = None
        st.session_state.mindmap_code = None
        st.session_state.flashcards_csv = None
        st.success("Memory Purged!")
        st.rerun()

    with st.expander("💬 Session History Logs"):
        if st.session_state.chat_history:
            for msg in st.session_state.chat_history:
                sender = "User" if msg["role"] == "user" else "Tutor"
                st.caption(f"**{sender}:** {msg['content'][:80]}...")
        else:
            st.caption("No chat log entries yet.")

    with st.expander("📊 Grade Report Cards"):
        if st.session_state.assessment_evaluation:
            st.caption(st.session_state.assessment_evaluation[:200] + "...")
        else:
            st.caption("No saved report cards.")

# --- MAIN INTERFACE HEADER (Text Only, Proper Case) ---
st.markdown("<h1 style='margin:0; padding-bottom: 4px;'>✨StudySphere</h1>", unsafe_allow_html=True)
st.caption(f"SYSTEM_MODE // {study_mode} | RIGOR // {academic_level.upper()}")
st.divider()

if not API_KEY:
    st.warning("🔒 Security Notice: Gemini API Key is missing. Please configure `.streamlit/secrets.toml` or set the environment variable.")
else:
    client = None
    try:
        client = genai.Client(api_key=API_KEY)
    except Exception as e:
        st.error(f"Initialization Error: {e}")

    # ==========================================
    # MODE 1: PRACTICE ASSESSMENT & TIMED QUIZ
    # ==========================================
    if study_mode == "📋 Practice Assessment & Evaluation":
        st.subheader("📝 Interactive Assessment & Timed Quiz Mode")
        st.write("")
        
        # --- SEPARATE BOXES FOR ASSESSMENT TYPE & DURATION ---
        col_type, col_timer = st.columns([1, 1], gap="medium")
        
        with col_type:
            with st.container(border=True):
                st.markdown("### 📑 Assessment Type")
                assessment_type = st.selectbox(
                    "Choose Format:",
                    [
                        "Mixed Exam (MCQs & Written)",
                        "Multiple Choice Questions (MCQ Only)",
                        "Written & Short Answer Assessment",
                        "Coding & Analytical Problem Solving"
                    ],
                    key="assessment_type_select"
                )
            
        with col_timer:
            with st.container(border=True):
                st.markdown("### ⏱️ Exam Duration")
                timer_options = {
                    "10 Mins": 10 * 60,
                    "30 Mins": 30 * 60,
                    "1 Hr": 60 * 60,
                    "3 Hrs": 180 * 60
                }
                
                selected_timer_label = st.segmented_control(
                    "Exam Duration",
                    options=list(timer_options.keys()),
                    default="30 Mins",
                    key="timer_control",
                    label_visibility="collapsed"
                )
                quiz_timer_sec = timer_options.get(selected_timer_label, 30 * 60)
        
        st.divider()
        
        if st.button("⚡ Generate Assessment Paper", use_container_width=True):
            if client:
                with st.spinner(f"Compiling {assessment_type} paper..."):
                    try:
                        if assessment_type == "Multiple Choice Questions (MCQ Only)":
                            prompt_format = (
                                "Generate exactly 4 MCQ questions.\n"
                                "Output strictly JSON array of objects: [{\"type\": \"mcq\", \"question\": \"...\", \"options\": [\"A)...\", \"B)...\", \"C)...\", \"D)...\"]}]"
                            )
                        elif assessment_type == "Written & Short Answer Assessment":
                            prompt_format = (
                                "Generate exactly 3 short-answer written conceptual questions.\n"
                                "Output strictly JSON array of objects: [{\"type\": \"written\", \"question\": \"...\"}]"
                            )
                        elif assessment_type == "Coding & Analytical Problem Solving":
                            prompt_format = (
                                "Generate exactly 3 technical problem-solving or coding logic tasks.\n"
                                "Output strictly JSON array of objects: [{\"type\": \"written\", \"question\": \"...\"}]"
                            )
                        else:  # Mixed Exam
                            prompt_format = (
                                "Generate 2 MCQs and 2 short written questions.\n"
                                "Output strictly JSON array of objects with keys 'type', 'question', and optional 'options'."
                            )

                        prompt = (
                            f"Generate an academic assessment paper for a {academic_level} level.\n"
                            f"Assessment Style: {assessment_type}.\n{prompt_format}"
                        )
                        
                        context_data = f"DOCUMENT MATERIAL:\n{document_text}" if document_text else "General academic topics."
                        response = client.models.generate_content(
                            model='gemini-2.5-flash',
                            contents=[f"Context: {context_data[:15000]}\n\nTask: {prompt}"],
                            config=types.GenerateContentConfig(response_mime_type="application/json")
                        )
                        st.session_state.assessment_questions = json.loads(response.text)
                        st.session_state.assessment_evaluation = None
                        st.session_state.start_time = time.time()
                        st.rerun()
                    except Exception as e:
                        st.error(f"Execution Error: {e}")

        if st.session_state.assessment_questions:
            st.divider()
            
            # Timed Countdown Display
            if "start_time" in st.session_state:
                elapsed = time.time() - st.session_state.start_time
                remaining = max(0, int(quiz_timer_sec - elapsed))
                hours, remainder = divmod(remaining, 3600)
                mins, secs = divmod(remainder, 60)
                
                timer_str = f"{hours:02d}:{mins:02d}:{secs:02d}" if hours > 0 else f"{mins:02d}:{secs:02d}"
                st.metric("⏳ Time Remaining", timer_str)
                if remaining == 0:
                    st.warning("⚠️ Time is up! Submit your answers now for evaluation.")

            st.subheader("✍️ Fill Answer Sheet")
            student_answers = []
            for i, q_data in enumerate(st.session_state.assessment_questions):
                with st.container(border=True):
                    st.markdown(f"### Question {i+1}")
                    st.markdown(f"**{q_data['question']}**")
                    
                    if q_data.get("type") == "mcq":
                        selected_option = st.radio("Select Option:", options=q_data["options"], key=f"mcq_{i}")
                        student_answers.append({"question": q_data["question"], "answer": selected_option, "type": "mcq"})
                    else:
                        ans_text = st.text_area("Your Answer:", key=f"written_{i}", placeholder="Type your response or code solution here...")
                        student_answers.append({"question": q_data["question"], "answer": ans_text, "type": "written"})
                
            if st.button("🚀 Grade My Answers", use_container_width=True):
                if client:
                    with st.spinner("Grading choices & written responses..."):
                        try:
                            eval_prompt = f"You are an examiner for a {academic_level} course. Evaluate this exam sheet thoroughly.\n\n"
                            for idx, item in enumerate(student_answers):
                                eval_prompt += f"Question {idx+1} ({item['type'].upper()}): {item['question']}\n"
                                eval_prompt += f"Student Answer {idx+1}: {item['answer']}\n\n"
                                
                            eval_prompt += (
                                "Provide an Evaluation Report Card with:\n"
                                "1. Verification of MCQ choices or code/written solutions.\n"
                                "2. Score out of 10 for each question.\n"
                                "3. Overall total score percentage and letter grade."
                            )
                            
                            eval_response = client.models.generate_content(
                                model='gemini-2.5-flash',
                                contents=[eval_prompt]
                            )
                            
                            st.session_state.assessment_evaluation = eval_response.text
                            st.session_state.questions_answered += len(student_answers)
                            save_local_settings(st.session_state.username, st.session_state.questions_answered, st.session_state.study_time_min)
                            st.rerun()
                        except Exception as e:
                            st.error(f"Evaluation Error: {e}")
                        
        if st.session_state.assessment_evaluation:
            st.divider()
            with st.container(border=True):
                st.subheader("📊 Exam Report Card")
                st.markdown(st.session_state.assessment_evaluation)
                
                st.download_button(
                    label="📥 Download Grade Report",
                    data=st.session_state.assessment_evaluation,
                    file_name="evaluation_report.txt",
                    mime="text/plain"
                )

    # ==========================================
    # MODE 2: FLASHCARD MAKER WITH ANKI CSV EXPORT
    # ==========================================
    elif study_mode == "⚡ Flashcard Maker (Anki Export)":
        st.subheader("⚡ Flashcard Deck Builder")
        topic_input = st.text_input("Enter Topic or Concept for Flashcards:", placeholder="e.g., Quantum Mechanics Principles")
        
        if st.button("⚡ Generate Anki-Ready Flashcard Deck", use_container_width=True):
            if client and (topic_input or document_text):
                with st.spinner("Generating flashcard pairs..."):
                    try:
                        prompt = (
                            f"Generate 5 high-yield flashcard Q&A pairs for a {academic_level} student based on the context/topic.\n"
                            "Format output strictly as a JSON array of objects with keys: 'front' and 'back'."
                        )
                        ctx = f"Topic: {topic_input}\nDocument Text: {document_text[:10000]}"
                        response = client.models.generate_content(
                            model='gemini-2.5-flash',
                            contents=[f"{ctx}\n\n{prompt}"],
                            config=types.GenerateContentConfig(response_mime_type="application/json")
                        )
                        cards = json.loads(response.text)
                        
                        csv_lines = ["Front,Back"]
                        for card in cards:
                            f_txt = card['front'].replace('"', '""')
                            b_txt = card['back'].replace('"', '""')
                            csv_lines.append(f'"{f_txt}","{b_txt}"')
                        
                        st.session_state.flashcards_csv = "\n".join(csv_lines)
                        st.session_state.flashcard_data = cards
                    except Exception as e:
                        st.error(f"Flashcard Generation Error: {e}")

        if "flashcard_data" in st.session_state and st.session_state.flashcard_data:
            st.divider()
            for idx, c in enumerate(st.session_state.flashcard_data):
                with st.expander(f"🎴 Flashcard {idx+1}: {c['front']}"):
                    st.markdown(f"**Answer:** {c['back']}")
            
            st.download_button(
                label="📥 Export Flashcards for Anki / Quizlet (.CSV)",
                data=st.session_state.flashcards_csv,
                file_name="anki_flashcards.csv",
                mime="text/csv"
            )

    # ==========================================
    # MODE 3: MIND MAP NODE VISUALIZER
    # ==========================================
    elif study_mode == "🧠 Mind Map Node Visualizer":
        st.subheader("🧠 Interactive Concept Mind Map Generator")
        mindmap_topic = st.text_input("Enter Concept to Map out:", placeholder="e.g., Neural Network Architecture")
        
        if st.button("⚡ Generate Mind Map Diagram", use_container_width=True):
            if client and (mindmap_topic or document_text):
                with st.spinner("Mapping core nodes and relationships..."):
                    try:
                        prompt = (
                            "Create a Graphviz DOT language graph representing the core concepts and sub-topics.\n"
                            "Rules:\n"
                            "- Return ONLY valid DOT language syntax starting with 'digraph G {' and ending with '}'.\n"
                            "- Do not include markdown code block backticks.\n"
                            "- Use dark mode neon hex colors like bgcolor='#0a050d', node fillcolor='#140b21', color='#d946ef', fontcolor='#ffffff'."
                        )
                        ctx = f"Topic: {mindmap_topic}\nMaterial: {document_text[:8000]}"
                        response = client.models.generate_content(
                            model='gemini-2.5-flash',
                            contents=[f"{ctx}\n\n{prompt}"]
                        )
                        dot_code = response.text.replace("```dot", "").replace("```", "").strip()
                        st.session_state.mindmap_code = dot_code
                    except Exception as e:
                        st.error(f"Mind Map Error: {e}")

        if st.session_state.mindmap_code:
            st.divider()
            try:
                graph = graphviz.Source(st.session_state.mindmap_code)
                st.graphviz_chart(st.session_state.mindmap_code)
            except Exception as e:
                st.error(f"Graphviz Rendering Error: {e}")

    # ==========================================
    # MODE 4: VISION NOTES & HANDWRITTEN READER
    # ==========================================
    elif study_mode == "👁️ Vision Notes & Formula Reader":
        st.subheader("👁️ Optical Note & Formula Breakdown")
        image_file = st.file_uploader("Upload Image of Notes, Whiteboards, or Math Equations", type=["png", "jpg", "jpeg"])
        
        if image_file and st.button("⚡ Analyze & Explain Note Image", use_container_width=True):
            if client:
                with st.spinner("Decoding image content with Gemini Vision..."):
                    try:
                        img_bytes = image_file.read()
                        image_part = types.Part.from_bytes(data=img_bytes, mime_type=image_file.type)
                        
                        vision_prompt = (
                            f"Transcribe and thoroughly explain all handwritten notes, formulas, or diagrams in this image "
                            f"for a {academic_level} student. Break down step-by-step solutions if math is present."
                        )
                        
                        response = client.models.generate_content(
                            model='gemini-2.5-flash',
                            contents=[image_part, vision_prompt]
                        )
                        
                        st.divider()
                        st.markdown("### 🔍 Image Breakdown & Solution")
                        st.markdown(response.text)
                        
                        # Audio Reader Synthesis
                        tts = gTTS(text=response.text[:500], lang='en')
                        speech_buffer = io.BytesIO()
                        tts.write_to_fp(speech_buffer)
                        st.audio(speech_buffer.getvalue(), format="audio/mp3")
                        
                    except Exception as e:
                        st.error(f"Vision Processing Error: {e}")

    # ==========================================
    # MODE 5: LECTURE MODE & AUDIO TTS CHAT
    # ==========================================
    else:
        base_instruction = (
            f"You are StudySphere, an expert engaging tutor. "
            f"Tutor the student at a {academic_level} level. Provide clear analogies and step-by-step points."
        )
        if document_text:
            base_instruction += f"\n\nCONTEXT FROM UPLOADED DOCUMENTS:\n{document_text[:15000]}"

        # Render Chat History
        for message in st.session_state.chat_history:
            avatar_icon = "🧑‍🎓" if message["role"] == "user" else "⚡"
            with st.chat_message(message["role"], avatar=avatar_icon):
                st.markdown(message["content"])

        if user_prompt := st.chat_input("Ask a question, enter a topic, or request an audio explanation..."):
            st.session_state.chat_history.append({"role": "user", "content": user_prompt})
            
            with st.chat_message("user", avatar="🧑‍🎓"):
                st.markdown(user_prompt)

            with st.chat_message("model", avatar="⚡"):
                message_placeholder = st.empty()
                if not client:
                    st.error("API Client missing or uninitialized.")
                else:
                    try:
                        formatted_contents = []
                        for m in st.session_state.chat_history[-6:]:
                            formatted_contents.append(
                                types.Content(
                                    role=m["role"],
                                    parts=[types.Part.from_text(text=m["content"])]
                                )
                            )
                        
                        response = client.models.generate_content(
                            model='gemini-2.5-flash',
                            contents=formatted_contents,
                            config=types.GenerateContentConfig(system_instruction=base_instruction)
                        )
                        
                        tutor_response = response.text
                        message_placeholder.markdown(tutor_response)
                        st.session_state.chat_history.append({"role": "model", "content": tutor_response})
                        
                        # Smart Text-To-Speech Audio Synthesis
                        tts = gTTS(text=tutor_response[:400], lang='en')
                        audio_fp = io.BytesIO()
                        tts.write_to_fp(audio_fp)
                        st.audio(audio_fp.getvalue(), format="audio/mp3")
                        
                    except Exception as e:
                        message_placeholder.error(f"API Error: {e}")

        if st.session_state.chat_history:
            chat_download_string = ""
            for msg in st.session_state.chat_history:
                role_label = "Student" if msg["role"] == "user" else "Tutor"
                chat_download_string += f"[{role_label}]: {msg['content']}\n\n"
            
            st.divider()
            st.download_button(
                label="📥 Export Chat Logs",
                data=chat_download_string,
                file_name="chat_session.txt",
                mime="text/plain"
            )