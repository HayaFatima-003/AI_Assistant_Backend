from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
from docx import Document
import os
from google import genai


# =========================================================
# GEMINI
# =========================================================

client = genai.Client(
    api_key=os.getenv("GEMINI_API_KEY")
)


# =========================================================
# LOAD WORD DOCUMENT
# =========================================================

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


# =========================================================
# LOAD EXCEL
# =========================================================

excel_data = pd.read_excel(
    "Permits Sample Data.xlsx"
)


# =========================================================
# WORD SECTION SEARCH
# =========================================================

def search_sections(question):

    question_lower = question.lower()

    keywords = {

        "HSD Truck Loading": [
            "hsd",
            "truck",
            "loading",
            "load",
            "diesel",
            "tank truck",
            "ppe",
            "hazard",
            "spill",
            "earthing"
        ],

        "HSD Decanting": [
            "decant",
            "decanting",
            "transfer",
            "tanker",
            "tote",
            "drum",
            "ppe",
            "hazard",
            "spill"
        ],

        "AOPS Testing": [
            "aops",
            "overfill",
            "alarm",
            "trip",
            "interlock",
            "shutdown",
            "testing"
        ],

        "Permit Requirements": [
            "permit",
            "ptw",
            "hot work",
            "cold work",
            "confined space",
            "excavation",
            "electrical",
            "lifting",
            "loto"
        ],

        "Basic HSE": [
            "hse",
            "ppe",
            "toolbox",
            "housekeeping",
            "spill",
            "incident",
            "near miss",
            "safety",
            "hazard"
        ],

        "Equipment Descriptions": [
            "equipment",
            "pump",
            "valve",
            "tank",
            "loading arm",
            "hose",
            "meter",
            "filter",
            "earthing"
        ],

        "Emergency Procedures": [
            "emergency",
            "fire",
            "injury",
            "evacuation",
            "alarm",
            "spill response",
            "muster"
        ]
    }


    scores = {}


    for section, words in keywords.items():

        score = 0

        for word in words:

            if word in question_lower:
                score += 1

        scores[section] = score


    best_section = max(
        scores,
        key=scores.get
    )


    if scores[best_section] == 0:

        return None, None


    # IMPORTANT:
    # Send the COMPLETE relevant section to Gemini.
    # Gemini will extract the exact answer.

    content = sections[best_section]


    return best_section, content


# =========================================================
# EXCEL SEARCH
# =========================================================

def search_excel(question):

    question_lower = question.lower()


    if "hot work" in question_lower:

        result = excel_data[
            excel_data["Permit Type"]
            .astype(str)
            .str.lower()
            .str.contains(
                "hot work",
                na=False
            )
        ]


    elif "cold work" in question_lower:

        result = excel_data[
            excel_data["Permit Type"]
            .astype(str)
            .str.lower()
            .str.contains(
                "cold work",
                na=False
            )
        ]


    elif "icc" in question_lower:

        result = excel_data[
            excel_data["Certificate"]
            .fillna("")
            .astype(str)
            .str.lower()
            .str.contains(
                "icc",
                na=False
            )
        ]


    elif "vec" in question_lower:

        result = excel_data[
            excel_data["Certificate"]
            .fillna("")
            .astype(str)
            .str.lower()
            .str.contains(
                "vec",
                na=False
            )
        ]


    elif "loading bay" in question_lower:

        result = excel_data[
            excel_data["Area"]
            .astype(str)
            .str.lower()
            .str.contains(
                "loading bay",
                na=False
            )
        ]


    elif "workshop" in question_lower:

        result = excel_data[
            excel_data["Area"]
            .astype(str)
            .str.lower()
            .str.contains(
                "workshop",
                na=False
            )
        ]


    elif "process area" in question_lower:

        result = excel_data[
            excel_data["Area"]
            .astype(str)
            .str.lower()
            .str.contains(
                "process area",
                na=False
            )
        ]


    else:

        return None


    return result


# =========================================================
# GEMINI ANSWER GENERATION
# =========================================================

def generate_ai_answer(
    question,
    context,
    source
):

    prompt = f"""
You are the KHT AI Assistant.

You are answering questions about a fictional/sample
hydrocarbon terminal training knowledge base.

Use ONLY the information provided in the context.

Rules:

1. Do not invent facts.
2. Do not add information that is not present
   in the context.
3. Answer the user's exact question directly.
4. Keep the answer concise and professional.
5. Use bullet points when appropriate.
6. Do not repeat the entire procedure unless
   the user specifically asks for the full procedure.
7. If the requested information is not present
   in the context, say:

"I could not find enough information in the
KHT knowledge base."

8. Treat all information as training/sample
   information and not as live operational instructions.

Source:
{source}

Knowledge Base Context:
{context}

User Question:
{question}
"""


    response = client.models.generate_content(

        model="gemini-3.8-flash",

        contents=prompt
    )


    return response.text


# =========================================================
# MAIN QUESTION ROUTER
# =========================================================

def ask_backend(question):

    question_lower = question.lower()


    # -----------------------------------------------------
    # EXCEL QUESTIONS
    # -----------------------------------------------------

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


    # -----------------------------------------------------
    # WORD QUESTIONS
    # -----------------------------------------------------

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
        "decanting",
        "earthing"

    ]


    # =====================================================
    # EXCEL ROUTING
    # =====================================================

    if any(
        word in question_lower
        for word in excel_words
    ):

        result = search_excel(question)


        if result is not None:

            return {

                "source": "Excel",

                "records_found": len(result),

                "data": result.head(10).to_dict(
                    orient="records"
                )

            }


    # =====================================================
    # WORD ROUTING
    # =====================================================

    if any(
        word in question_lower
        for word in word_words
    ):

        section, content = search_sections(
            question
        )


        if section:

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


    # =====================================================
    # NOTHING FOUND
    # =====================================================

    return {

        "source": "Unknown",

        "message": (
            "I could not find relevant information "
            "in the KHT knowledge base."
        )

    }


# =========================================================
# FASTAPI
# =========================================================

app = FastAPI(

    title="KHT AI Assistant",

    version="1.0"

)


# =========================================================
# CORS
# =========================================================

app.add_middleware(

    CORSMiddleware,

    allow_origins=["*"],

    allow_credentials=True,

    allow_methods=["*"],

    allow_headers=["*"]

)


# =========================================================
# HOME
# =========================================================

@app.get("/")
def home():

    return {

        "message":
        "KHT AI Assistant backend is running!"

    }


# =========================================================
# ASK ENDPOINT
# =========================================================

@app.get("/ask")
def ask(question: str):

    return ask_backend(question)
