
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
from docx import Document
import os
from google import genai


client = genai.Client(
    api_key=os.getenv("GEMINI_API_KEY")
)

# =========================
# LOAD WORD KNOWLEDGE BASE
# =========================

doc = Document("KHT Operation knowledge sample data.docx")

sections = {
    "HSD Truck Loading": "",
    "HSD Decanting": "",
    "AOPS Testing": "",
    "Permit Requirements": "",
    "Basic HSE": "",
    "Equipment Descriptions": "",
    "Emergency Procedures": ""
}

current_section = None

for paragraph in doc.paragraphs:
    text = paragraph.text.strip()

    if "Operation 1: HSD Truck Loading" in text:
        current_section = "HSD Truck Loading"
    elif "Operation 2: HSD Decanting" in text:
        current_section = "HSD Decanting"
    elif "Operation 3: AOPS Testing" in text:
        current_section = "AOPS Testing"
    elif "Operation 4: Permit Requirements" in text:
        current_section = "Permit Requirements"
    elif "Operation 5: Basic HSE" in text:
        current_section = "Basic HSE"
    elif "Operation 6: Equipment Descriptions" in text:
        current_section = "Equipment Descriptions"
    elif "Operation 7: Emergency Procedures" in text:
        current_section = "Emergency Procedures"

    if current_section:
        sections[current_section] += text + "\n"


# =========================
# LOAD EXCEL DATA
# =========================

excel_data = pd.read_excel("Permits Sample Data.xlsx")


# =========================
# WORD SEARCH
# =========================


# =========================
# WORD SEARCH
# =========================

import re


def search_sections(question):

    question_lower = question.lower()

    keywords = {
        "HSD Truck Loading": [
            "hsd", "truck", "loading", "load",
            "diesel", "tank truck", "ppe",
            "hazard", "spill", "earthing"
        ],

        "HSD Decanting": [
            "decant", "decanting", "transfer",
            "tanker", "tote", "drum",
            "ppe", "hazard", "spill"
        ],

        "AOPS Testing": [
            "aops", "overfill", "alarm",
            "trip", "interlock", "shutdown",
            "testing"
        ],

        "Permit Requirements": [
            "permit", "ptw", "hot work",
            "cold work", "confined space",
            "excavation", "electrical",
            "lifting", "loto"
        ],

        "Basic HSE": [
            "hse", "ppe", "toolbox",
            "housekeeping", "spill",
            "incident", "near miss",
            "safety", "hazard"
        ],

        "Equipment Descriptions": [
            "equipment", "pump", "valve",
            "tank", "loading arm",
            "hose", "meter", "filter",
            "earthing"
        ],

        "Emergency Procedures": [
            "emergency", "fire", "injury",
            "evacuation", "alarm",
            "spill response", "muster"
        ]
    }

    # -------------------------
    # Find best section
    # -------------------------

    scores = {}

    for section, words in keywords.items():

        score = 0

        for word in words:
            if word in question_lower:
                score += 1

        scores[section] = score

    best_section = max(scores, key=scores.get)

    if scores[best_section] == 0:
        return None, None

    content = sections[best_section]

    # -------------------------
    # Break content into sentences
    # -------------------------

    sentences = re.split(
        r'(?<=[.!?])\s+',
        content.strip()
    )

    # -------------------------
    # Remove common words
    # -------------------------

    stop_words = {
        "what", "is", "are", "the", "a", "an",
        "for", "of", "to", "in", "on", "and",
        "or", "how", "does", "do", "can",
        "should", "be", "required", "tell",
        "me", "please", "about"
    }

    question_words = set(
        re.findall(r'\b[a-zA-Z0-9]+\b', question_lower)
    )

    question_words -= stop_words

    # -------------------------
    # Score each sentence
    # -------------------------

    scored_sentences = []

    for sentence in sentences:

        sentence_lower = sentence.lower()

        sentence_words = set(
            re.findall(
                r'\b[a-zA-Z0-9]+\b',
                sentence_lower
            )
        )

        overlap = question_words.intersection(
            sentence_words
        )

        score = len(overlap)

        if score > 0:
            scored_sentences.append(
                (score, sentence.strip())
            )

    # -------------------------
    # Sort most relevant first
    # -------------------------

    scored_sentences.sort(
        key=lambda x: x[0],
        reverse=True
    )

    # -------------------------
    # Return concise answer
    # -------------------------

    if scored_sentences:

        top_sentences = [
            sentence
            for score, sentence
            in scored_sentences[:3]
        ]

        answer = " ".join(top_sentences)

    else:

        answer = content

    return best_section, answer
  


# =========================
# EXCEL SEARCH
# =========================

def search_excel(question):

    question = question.lower()

    if "hot work" in question:

        result = excel_data[
            excel_data["Permit Type"]
            .astype(str)
            .str.lower()
            .str.contains("hot work", na=False)
        ]

    elif "cold work" in question:

        result = excel_data[
            excel_data["Permit Type"]
            .astype(str)
            .str.lower()
            .str.contains("cold work", na=False)
        ]

    elif "icc" in question:

        result = excel_data[
            excel_data["Certificate"]
            .fillna("")
            .astype(str)
            .str.lower()
            .str.contains("icc", na=False)
        ]

    elif "vec" in question:

        result = excel_data[
            excel_data["Certificate"]
            .fillna("")
            .astype(str)
            .str.lower()
            .str.contains("vec", na=False)
        ]

    elif "loading bay" in question:

        result = excel_data[
            excel_data["Area"]
            .astype(str)
            .str.lower()
            .str.contains("loading bay", na=False)
        ]

    elif "workshop" in question:

        result = excel_data[
            excel_data["Area"]
            .astype(str)
            .str.lower()
            .str.contains("workshop", na=False)
        ]

    elif "process area" in question:

        result = excel_data[
            excel_data["Area"]
            .astype(str)
            .str.lower()
            .str.contains("process area", na=False)
        ]

    else:
        return None

    return result
# =========================
# GEMINI AI ANSWER
# =========================

def generate_ai_answer(question, context, source):

    prompt = f"""
You are the KHT AI Assistant.

Answer the user's question using ONLY the information provided
in the knowledge-base context below.

Do not invent facts.
Do not add information that is not present in the context.
If the context does not contain enough information, say:
"I could not find enough information in the KHT knowledge base."

Give a concise, professional answer.
Use bullet points when they make the answer clearer.

Source: {source}

Knowledge-base context:
{context}

User question:
{question}
"""

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt
    )

    return response.text

# =========================
# MAIN BACKEND LOGIC
# =========================

def ask_backend(question):

    question_lower = question.lower()

    excel_words = [
        "how many",
        "count",
        "permit number",
        "permits",
        "records",
        "certificate",
        "supervisor",
        "issued",
        "closing date",
        "area"
    ]

    word_words = [
        "procedure",
        "how to",
        "steps",
        "required",
        "ppe",
        "hazard",
        "safety",
        "emergency",
        "equipment",
        "aops",
        "loading",
        "decanting"
    ]

    if any(word in question_lower for word in excel_words):

        result = search_excel(question)

        if result is not None:

            return {
                "source": "Excel",
                "records_found": len(result),
                "data": result.head(10).to_dict(
                    orient="records"
                )
            }

    if any(word in question_lower for word in word_words):

        section, content = search_sections(question)

        if section:

            return {
                "source": "Word",
                "section": section,
                "content": content
            }

  answer = generate_ai_answer(
    question,
    content,
    f"Word Knowledge Base - {section}"
)

return {
    "source": "Word",
    "section": section,
    "content": answer
}

# =========================
# FASTAPI APP
# =========================

app = FastAPI(
    title="KHT AI Assistant",
    version="1.0"
)


# =========================
# CORS
# =========================

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"]
)


# =========================
# ROUTES
# =========================

@app.get("/")
def home():

    return {
        "message": "KHT AI Assistant backend is running!"
    }


@app.get("/ask")
def ask(question: str):

    return ask_backend(question)
