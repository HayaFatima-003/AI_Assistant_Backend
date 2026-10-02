# KHT AI Assistant

AI knowledge assistant for Karachi Hydrocarbon Terminal.

## Architecture

Frontend
→ FastAPI Backend
→ Gemini API
→ KHT Word + Excel Knowledge

## Environment Variable

Set:

GEMINI_API_KEY=your_key_here

## Start

uvicorn main:app --host 0.0.0.0 --port $PORT

## Endpoints

GET /

GET /health

POST /ask
