import os
import re
from pathlib import Path
from contextlib import asynccontextmanager

import pandas as pd
from docx import Document
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from google import genai


# ============================================================
# PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

DOCX_PATH = DATA_DIR / "KHT Operation knowledge sample data.docx"
XLSX_PATH = DATA_DIR / "Permits Sample Data.xlsx"


# ============================================================
# GEMINI SETUP
# ============================================================

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

MODEL_NAME = "gemini-3.5-flash-lite"

client = None

if GEMINI_API_KEY:
    client = genai.Client(api_key=GEMINI_API_KEY)
else:
    print("WARNING: GEMINI_API_KEY is not configured.")


# ============================================================
# DATA
# ============================================================

word_sections = []
permit_df = None


# ============================================================
# NORMALIZE TEXT
# ============================================================

def normalize(text):

    if text is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(text).lower().strip()
    )


# ============================================================
# LOAD WORD DOCUMENT
# ============================================================

def load_word_document():

    global word_sections

    word_sections = []

    if not DOCX_PATH.exists():

        print(f"ERROR: Word file not found: {DOCX_PATH}")

        return

    try:

        document = Document(DOCX_PATH)

        current_title = "General KHT Knowledge"
        current_text = []

        for paragraph in document.paragraphs:

            text = paragraph.text.strip()

            if not text:
                continue

            if (
                text.startswith("Operation ")
                or text.startswith("General ")
                or text.startswith("Emergency ")
                or text.startswith("Basic ")
                or text.startswith("Equipment ")
                or text.startswith("Permit ")
            ):

                if current_text:

                    word_sections.append({
                        "title": current_title,
                        "content": "\n".join(current_text)
                    })

                current_title = text
                current_text = []

            else:

                current_text.append(text)

        if current_text:

            word_sections.append({
                "title": current_title,
                "content": "\n".join(current_text)
            })

        print(
            f"Word sections loaded: {len(word_sections)}"
        )

    except Exception as e:

        print(
            f"ERROR loading Word document: {type(e).__name__}: {e}"
        )

        word_sections = []


# ============================================================
# LOAD EXCEL
# ============================================================

def load_excel():

    global permit_df

    permit_df = None

    if not XLSX_PATH.exists():

        print(f"ERROR: Excel file not found: {XLSX_PATH}")

        return

    try:

        permit_df = pd.read_excel(XLSX_PATH)

        print(
            f"Excel rows loaded: {len(permit_df)}"
        )

        print(
            f"Excel columns: {list(permit_df.columns)}"
        )

    except Exception as e:

        print(
            f"ERROR loading Excel: {type(e).__name__}: {e}"
        )

        permit_df = None


# ============================================================
# SEARCH WORD DOCUMENT
# ============================================================

def search_word_sections(question, top_k=3):

    question = normalize(question)

    keyword_groups = {

        "ppe": [
            "ppe",
            "personal protective equipment",
            "protective equipment",
            "helmet",
            "gloves",
            "safety shoes",
            "goggles",
            "coverall"
        ],

        "loading": [
            "loading",
            "truck loading",
            "hsd loading",
            "truck",
            "loading procedure"
        ],

        "decanting": [
            "decanting",
            "decant",
            "hsd decanting"
        ],

        "aops": [
            "aops",
            "testing",
            "aops testing"
        ],

        "permit": [
            "permit",
            "ptw",
            "hot work",
            "cold work",
            "icc",
            "vec",
            "certificate"
        ],

        "hse": [
            "hse",
            "safety",
            "hazard",
            "ppe"
        ],

        "equipment": [
            "equipment",
            "pump",
            "valve",
            "tank",
            "pipeline"
        ],

        "emergency": [
            "emergency",
            "fire",
            "spill",
            "leak",
            "incident"
        ]
    }

    scores = []

    question_words = set(
        word
        for word in question.split()
        if len(word) > 2
    )

    for section in word_sections:

        title = normalize(section["title"])
        content = normalize(section["content"])

        score = 0

        for keywords in keyword_groups.values():

            for keyword in keywords:

                if keyword in question:

                    if keyword in title:
                        score += 8

                    if keyword in content:
                        score += 3

        content_words = set(content.split())

        overlap = question_words.intersection(
            content_words
        )

        score += len(overlap)

        if score > 0:

            scores.append(
                (
                    score,
                    section["title"],
                    section["content"]
                )
            )

    scores.sort(
        reverse=True,
        key=lambda x: x[0]
    )

    return scores[:top_k]


# ============================================================
# SEARCH EXCEL
# ============================================================

def search_excel(question, top_k=8):

    if permit_df is None:
        return []

    question = normalize(question)

    results = []

    question_words = set(
        word
        for word in question.split()
        if len(word) > 2
    )

    important_terms = [
        "hot work",
        "cold work",
        "permit",
        "icc",
        "vec",
        "certificate",
        "loading",
        "decanting"
    ]

    for _, row in permit_df.iterrows():

        row_parts = []

        for column in permit_df.columns:

            value = row[column]

            if pd.notna(value):

                row_parts.append(
                    f"{column}: {value}"
                )

        row_text = " | ".join(row_parts)

        normalized_row = normalize(row_text)

        score = 0

        for word in question_words:

            if word in normalized_row:
                score += 1

        for term in important_terms:

            if (
                term in question
                and term in normalized_row
            ):

                score += 5

        if score > 0:

            results.append(
                (
                    score,
                    row_text
                )
            )

    results.sort(
        reverse=True,
        key=lambda x: x[0]
    )

    return [
        result[1]
        for result in results[:top_k]
    ]


# ============================================================
# BUILD KHT CONTEXT
# ============================================================

def build_context(question):

    word_results = search_word_sections(
        question,
        top_k=3
    )

    excel_results = search_excel(
        question,
        top_k=8
    )

    context_parts = []

    for score, title, content in word_results:

        context_parts.append(
            f"SOURCE: {title}\n{content}"
        )

    if excel_results:

        context_parts.append(
            "SOURCE: KHT Permit Database\n"
            + "\n".join(excel_results)
        )

    if not context_parts:

        return "No directly matching KHT information was found."

    return "\n\n".join(context_parts)


# ============================================================
# CLEAN TEXT
# ============================================================

def clean_text(text):

    return (
        str(text)
        .replace("\u2014", "-")
        .replace("\u2013", "-")
        .replace("\u2018", "'")
        .replace("\u2019", "'")
        .replace("\u201c", '"')
        .replace("\u201d", '"')
    )


# ============================================================
# ASK GEMINI
# ============================================================

def ask_gemini(question, context):

    if client is None:

        return (
            "Gemini API is not configured. "
            "Please check the backend environment variable."
        )

    question = clean_text(question)
    context = clean_text(context)

    prompt = f"""
You are the KHT AI Assistant for Karachi Hydrocarbon Terminal.

Answer the user's question using ONLY the KHT information
provided below.

Rules:

- Keep the answer concise and practical.
- Do not invent procedures, permit numbers, dates,
  certificates, equipment details, or safety requirements.
- If the information is not available, say so clearly.
- For simple questions, use short bullet points.
- Mention the relevant source at the end.
- This is a training/demo assistant.
- Approved site procedures must always be verified
  before operational use.

KHT INFORMATION:

{context}

USER QUESTION:

{question}

ANSWER:
"""

    prompt = clean_text(prompt)

    try:

        response = client.models.generate_content(
            model=MODEL_NAME,
            contents=prompt
        )

        if response.text:

            return response.text.strip()

        return "No answer was generated."

    except Exception as e:

        print(
            f"Gemini error: {type(e).__name__}: {e}"
        )

        return (
            "The KHT AI Assistant could not generate "
            "an answer right now."
        )


# ============================================================
# FASTAPI LIFESPAN
# ============================================================

@asynccontextmanager
async def lifespan(app: FastAPI):

    print("=" * 60)
    print("STARTING KHT AI ASSISTANT")
    print("=" * 60)

    print("Word file:")
    print(DOCX_PATH)

    print("Excel file:")
    print(XLSX_PATH)

    print("Gemini model:")
    print(MODEL_NAME)

    print(
        "Gemini API configured:",
        bool(GEMINI_API_KEY)
    )

    load_word_document()
    load_excel()

    print("=" * 60)
    print("KHT AI ASSISTANT READY")
    print("=" * 60)

    yield


# ============================================================
# FASTAPI APP
# ============================================================

app = FastAPI(
    title="KHT AI Assistant",
    description="KHT operational knowledge and permit assistant",
    version="1.0",
    lifespan=lifespan
)


# ============================================================
# CORS
# ============================================================

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"]
)


# ============================================================
# REQUEST MODEL
# ============================================================

class QuestionRequest(BaseModel):

    question: str


# ============================================================
# ROOT
# ============================================================

@app.get("/")
def root():

    return {
        "status": "online",
        "assistant": "KHT AI Assistant",
        "model": MODEL_NAME,
        "gemini_configured": bool(GEMINI_API_KEY),
        "word_sections": len(word_sections),
        "permit_rows": (
            len(permit_df)
            if permit_df is not None
            else 0
        )
    }


# ============================================================
# HEALTH
# ============================================================

@app.get("/health")
def health():

    return {
        "status": "healthy",
        "model": MODEL_NAME,
        "gemini_configured": bool(GEMINI_API_KEY),
        "word_sections": len(word_sections),
        "permit_rows": (
            len(permit_df)
            if permit_df is not None
            else 0
        )
    }


# ============================================================
# ASK
# ============================================================

@app.post("/ask")
def ask(request: QuestionRequest):

    question = request.question.strip()

    if not question:

        return {
            "answer": "Please enter a question.",
            "source": "KHT AI Assistant"
        }

    print(
        f"Question received: {question}"
    )

    context = build_context(question)

    answer = ask_gemini(
        question,
        context
    )

    return {
        "answer": answer,
        "source": "KHT Knowledge Base"
    }
