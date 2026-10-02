
import os
import re
from pathlib import Path

import pandas as pd
from docx import Document
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from google import genai
from google.genai import types


# ============================================================
# SETTINGS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

DOCX_PATH = DATA_DIR / "KHT Operation knowledge sample data.docx"
XLSX_PATH = DATA_DIR / "Permits Sample Data.xlsx"

MODEL_NAME = "gemini-3.8-flash"

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

if not GEMINI_API_KEY:
    raise RuntimeError(
        "GEMINI_API_KEY environment variable is missing."
    )

client = genai.Client(
    api_key=GEMINI_API_KEY
)


# ============================================================
# LOAD WORD KNOWLEDGE
# ============================================================

document = Document(str(DOCX_PATH))

word_sections = []

current_title = "General KHT Knowledge"
current_text = []

for paragraph in document.paragraphs:

    text = paragraph.text.strip()

    if not text:
        continue

    style_name = paragraph.style.name.lower()

    is_heading = (
        style_name.startswith("heading")
        or (
            len(text) < 100
            and text.isupper()
            and len(text.split()) <= 12
        )
    )

    if is_heading:

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


# ============================================================
# LOAD EXCEL
# ============================================================

permit_df = pd.read_excel(
    XLSX_PATH
)

permit_df.columns = [
    str(col).strip()
    for col in permit_df.columns
]

permit_df = permit_df.fillna("")


# ============================================================
# NORMALIZATION
# ============================================================

def normalize_text(text):

    text = str(text).lower()

    text = re.sub(
        r"[^a-z0-9\s\-\/]",
        " ",
        text
    )

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


# ============================================================
# WORD SEARCH
# ============================================================

def search_word_sections(question, top_k=3):

    q = normalize_text(question)

    q_words = set(q.split())

    scored = []

    keyword_groups = {

        "ppe": [
            "ppe",
            "personal protective",
            "helmet",
            "gloves",
            "goggles"
        ],

        "loading": [
            "loading",
            "truck loading",
            "hsd loading"
        ],

        "decanting": [
            "decanting",
            "decant"
        ],

        "permit": [
            "permit",
            "hot work",
            "cold work",
            "icc",
            "vec"
        ],

        "emergency": [
            "emergency",
            "fire",
            "spill",
            "evacuation"
        ],

        "aops": [
            "aops",
            "testing",
            "live testing"
        ]
    }

    for section in word_sections:

        title = section["title"]
        content = section["content"]

        combined = normalize_text(
            title + " " + content
        )

        score = 0

        if q in combined:
            score += 20

        section_words = set(combined.split())

        overlap = q_words.intersection(
            section_words
        )

        score += len(overlap) * 2

        for group_words in keyword_groups.values():

            q_has_group = any(
                word in q
                for word in group_words
            )

            section_has_group = any(
                word in combined
                for word in group_words
            )

            if q_has_group and section_has_group:
                score += 8

        if score > 0:

            scored.append({
                "score": score,
                "title": title,
                "content": content
            })

    scored.sort(
        key=lambda x: x["score"],
        reverse=True
    )

    return scored[:top_k]


# ============================================================
# EXCEL SEARCH
# ============================================================

def search_excel(question, top_k=8):

    q = normalize_text(question)

    matches = []

    for _, row in permit_df.iterrows():

        row_text = " ".join(
            str(value)
            for value in row.values
        )

        normalized_row = normalize_text(
            row_text
        )

        score = 0

        if q in normalized_row:
            score += 15

        q_words = set(q.split())
        row_words = set(
            normalized_row.split()
        )

        score += len(
            q_words.intersection(row_words)
        )

        if "hot work" in q and "hot work" in normalized_row:
            score += 10

        if "cold work" in q and "cold work" in normalized_row:
            score += 10

        if "icc" in q and "icc" in normalized_row:
            score += 10

        if "vec" in q and "vec" in normalized_row:
            score += 10

        if "permit" in q and "permit" in normalized_row:
            score += 3

        if score > 0:

            matches.append({
                "score": score,
                "row": row.astype(str).to_dict()
            })

    matches.sort(
        key=lambda x: x["score"],
        reverse=True
    )

    return matches[:top_k]


# ============================================================
# CONTEXT
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

    for result in word_results:

        context_parts.append(
            f"""
SOURCE TYPE: KHT Operation Knowledge
SECTION: {result['title']}

{result['content']}
"""
        )

    if excel_results:

        excel_text = []

        for result in excel_results:

            row = result["row"]

            formatted = " | ".join(
                f"{key}: {value}"
                for key, value in row.items()
                if str(value).strip()
            )

            excel_text.append(
                formatted
            )

        context_parts.append(
            """
SOURCE TYPE: KHT Permit Records
DATA:

""" + "\n".join(excel_text)
        )

    if not context_parts:

        return (
            "No directly relevant KHT information "
            "was found in the available knowledge base."
        )

    return "\n\n".join(context_parts)


# ============================================================
# SYSTEM INSTRUCTION
# ============================================================

SYSTEM_INSTRUCTION = """
You are KHT AI Assistant.

You are an internal knowledge assistant for Karachi Hydrocarbon Terminal.

Use ONLY the KHT knowledge and permit information supplied in the context.

Rules:

- Do not invent KHT procedures or requirements.
- If the answer is not in the context, say so.
- Keep answers concise and practical.
- Use bullets when appropriate.
- Do not expose API keys or internal code.
- Training/sample information must not be presented as approved site procedure.
- Safety-critical questions must be answered conservatively.
- Mention the relevant source section when useful.
"""


# ============================================================
# GEMINI FUNCTION
# ============================================================

def generate_answer(question):

    context = build_context(question)

    prompt = f"""
KHT USER QUESTION:
{question}

KHT KNOWLEDGE CONTEXT:
{context}

Answer the question using the supplied KHT context.

Keep the answer concise.

If the information is not available in the context, say:

"I couldn't find this information in the available KHT knowledge base."

At the end provide a short source label.
"""

    response = client.models.generate_content(
        model=MODEL_NAME,
        contents=prompt,
        config=types.GenerateContentConfig(
            system_instruction=SYSTEM_INSTRUCTION,
            temperature=0.2,
            max_output_tokens=500
        )
    )

    return response.text.strip()


# ============================================================
# FASTAPI
# ============================================================

app = FastAPI(
    title="KHT AI Assistant",
    version="1.0"
)


app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
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
        "model": MODEL_NAME
    }


# ============================================================
# HEALTH
# ============================================================

@app.get("/health")
def health():

    return {
        "status": "healthy",
        "word_sections": len(word_sections),
        "permit_records": len(permit_df),
        "model": MODEL_NAME
    }


# ============================================================
# ASK
# ============================================================

@app.post("/ask")
def ask(request: QuestionRequest):

    question = request.question.strip()

    if not question:

        raise HTTPException(
            status_code=400,
            detail="Question cannot be empty."
        )

    try:

        answer = generate_answer(
            question
        )

        word_results = search_word_sections(
            question,
            top_k=3
        )

        excel_results = search_excel(
            question,
            top_k=8
        )

        sources = []

        for result in word_results:

            sources.append(
                result["title"]
            )

        if excel_results:

            sources.append(
                "KHT Permit Records"
            )

        return {
            "answer": answer,
            "source": ", ".join(sources)
                if sources
                else "KHT Knowledge Base",
            "records_found": (
                len(word_results)
                + len(excel_results)
            )
        }

    except Exception as e:

        print("ERROR:", repr(e))

        raise HTTPException(
            status_code=500,
            detail="The KHT AI Assistant could not process the request."
        )
