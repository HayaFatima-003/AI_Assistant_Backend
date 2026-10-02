# ============================================================
# KHT AI ASSISTANT - COLAB TEST
# Gemini 3.5 Flash-Lite + KHT DOCX + Permit Excel
# ============================================================

from google import genai
from docx import Document
import pandas as pd
from pathlib import Path
import os
import re


# ============================================================
# 1. GEMINI SETUP
# ============================================================

api_key = input("Enter your NEW Gemini API key: ").strip()

# Remove accidental spaces/newlines only
api_key = api_key.strip()

client = genai.Client(api_key=api_key)

MODEL_NAME = "gemini-3.5-flash-lite"

print("\nGemini client created.")
print("Model:", MODEL_NAME)


# ============================================================
# 2. FILE PATHS
# ============================================================

BASE_DIR = Path("/content/KHT-AI-Backend")
DATA_DIR = BASE_DIR / "data"

DOCX_PATH = DATA_DIR / "KHT Operation knowledge sample data.docx"
XLSX_PATH = DATA_DIR / "Permits Sample Data.xlsx"

print("\nChecking files...")

print("DOCX:", DOCX_PATH)
print("Exists:", DOCX_PATH.exists())

print("Excel:", XLSX_PATH)
print("Exists:", XLSX_PATH.exists())


# ============================================================
# 3. LOAD WORD DOCUMENT
# ============================================================

word_sections = []


def load_word_document():

    global word_sections

    word_sections = []

    if not DOCX_PATH.exists():
        print("\nERROR: Word file not found.")
        return

    document = Document(DOCX_PATH)

    current_title = "General KHT Knowledge"
    current_text = []

    for paragraph in document.paragraphs:

        text = paragraph.text.strip()

        if not text:
            continue

        # Detect section headings
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

    # Final section
    if current_text:
        word_sections.append({
            "title": current_title,
            "content": "\n".join(current_text)
        })

    print("\nWord sections loaded:", len(word_sections))


# ============================================================
# 4. LOAD EXCEL
# ============================================================

permit_df = None


def load_excel():

    global permit_df

    if not XLSX_PATH.exists():
        print("\nERROR: Excel file not found.")
        return

    permit_df = pd.read_excel(XLSX_PATH)

    print("Excel rows loaded:", len(permit_df))
    print("Excel columns:", list(permit_df.columns))


# ============================================================
# 5. NORMALIZE TEXT
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
# 6. SEARCH WORD DOCUMENT
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
        word for word in question.split()
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

        overlap = question_words.intersection(content_words)

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
# 7. SEARCH EXCEL
# ============================================================

def search_excel(question, top_k=8):

    if permit_df is None:
        return []

    question = normalize(question)

    results = []

    question_words = set(
        word for word in question.split()
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

            if term in question and term in normalized_row:
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
# 8. BUILD KHT CONTEXT
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

    # Word results
    for score, title, content in word_results:

        context_parts.append(
            f"SOURCE: {title}\n{content}"
        )

    # Excel results
    if excel_results:

        context_parts.append(
            "SOURCE: KHT Permit Database\n"
            + "\n".join(excel_results)
        )

    return "\n\n".join(context_parts)


# ============================================================
# 9. ASK GEMINI
# ============================================================

def ask_gemini(question, context):

    prompt = f"""
You are the KHT AI Assistant for Karachi Hydrocarbon Terminal.

Answer the user's question using ONLY the KHT information
provided below.

Rules:
- Keep the answer concise and practical.
- Do not invent procedures, permit numbers, dates, certificates,
  equipment details, or safety requirements.
- If the information is not available, say so clearly.
- For simple questions, use short bullet points.
- Mention the relevant source at the end.
- This is a training/demo assistant. Approved site procedures
  must always be verified before operational use.

KHT INFORMATION:

{context}

USER QUESTION:

{question}

ANSWER:
"""

    # Remove problematic Unicode characters that can cause
    # header/request encoding issues.
    prompt = (
        prompt
        .replace("\u2014", "-")
        .replace("\u2013", "-")
        .replace("\u2018", "'")
        .replace("\u2019", "'")
        .replace("\u201c", '"')
        .replace("\u201d", '"')
    )

    response = client.models.generate_content(
        model=MODEL_NAME,
        contents=prompt
    )

    return response.text


# ============================================================
# 10. LOAD EVERYTHING
# ============================================================

print("\n" + "=" * 60)
print("LOADING KHT KNOWLEDGE")
print("=" * 60)

load_word_document()
load_excel()


# ============================================================
# 11. TEST GEMINI
# ============================================================

print("\n" + "=" * 60)
print("TESTING GEMINI")
print("=" * 60)

try:

    test_response = client.models.generate_content(
        model=MODEL_NAME,
        contents="Reply with exactly: KHT TEST OK"
    )

    print("Gemini:", test_response.text)

except Exception as e:

    print("Gemini test FAILED:")
    print(type(e).__name__)
    print(str(e))


# ============================================================
# 12. TEST KHT QUESTION
# ============================================================

print("\n" + "=" * 60)
print("TESTING KHT QUESTION")
print("=" * 60)

question = "What PPE is required for HSD truck loading?"

context = build_context(question)

print("\nContext length:", len(context))

if context:

    print("\nRelevant KHT context found:")
    print(context[:2500])

    print("\n" + "=" * 60)
    print("KHT AI ANSWER")
    print("=" * 60)

    try:

        answer = ask_gemini(
            question,
            context
        )

        print(answer)

    except Exception as e:

        print("\nGemini answer FAILED:")
        print(type(e).__name__)
        print(str(e))

else:

    print("\nWARNING: No relevant KHT context found.")


# ============================================================
# 13. FINAL STATUS
# ============================================================

print("\n" + "=" * 60)
print("FINAL STATUS")
print("=" * 60)

print("Gemini model:", MODEL_NAME)
print("Word sections:", len(word_sections))
print(
    "Excel rows:",
    len(permit_df) if permit_df is not None else 0
)
print("KHT backend logic test completed.")
