# 🎓 OmniStudy AI: University Academic & Research Suite

A Streamlit study workspace for university students. Upload lecture slides, papers or notes (or paste text), pick a model, and generate summaries, flashcards, graded practice exams, expanded lecture notes, research-paper critiques and term-project roadmaps.

## Features

| Tab | What it does |
|---|---|
| Summary & Study Guides | Executive summary, exam revision guide, or first-principles explainer |
| Flashcard Studio | 5–50 cards, expandable review, spaced-repetition drill, Markdown and Anki/Quizlet (TSV) export |
| Practice MCQ Exam | 3–25 questions, 3 difficulty levels, graded score card, rationale for every option |
| Slide-to-Notes | Expands sparse slide bullets into textbook-style chapters with formulas and exam essay prompts |
| Research Paper Analyzer | Research question, hypothesis, methodology, findings, critique, APA 7th and IEEE citations |
| Term Project Builder | Architecture, four-phase milestones, deliverables checklist, tech stack, risks |

Every output can be downloaded as Markdown.

**Input formats:** PDF, DOCX, PPTX (including speaker notes and tables), TXT, plus pasted text. Multiple files are combined.

## Models

Models are reached through OpenRouter using the standard `openai` SDK (`base_url="https://openrouter.ai/api/v1"`). Categories in the sidebar:

- **Frontier Academic & Reasoning:** Claude 3.5 Sonnet, GPT-4o, DeepSeek R1, Gemini
- **Open-Weights Powerhouses:** Llama 3.3 70B, Qwen 2.5 72B, Mistral Large / Nemo
- **Zero-Cost Free Auto-Router:** `openrouter/free`
- **Direct Google AI Studio:** Gemini via your own Google key

Model IDs on OpenRouter change often, especially `:free` ones. If a model stops working, paste a current ID into **Custom model ID** in the sidebar. By default the app falls back to `openrouter/free` when a model fails and tells you which model actually answered.

## Quick start (local)

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then paste an OpenRouter key into the sidebar, or create `.streamlit/secrets.toml`:

```toml
OPENROUTER_API_KEY = "sk-or-v1-..."
GOOGLE_API_KEY = "AIza..."   # optional
```

Do not commit `secrets.toml` to GitHub.

## Deploy

Step-by-step instructions for getting free keys, uploading to GitHub through the web UI, and deploying on Streamlit Community Cloud are in [`DEPLOYMENT_GUIDE.md`](DEPLOYMENT_GUIDE.md).

## Project structure

```
app.py                  The whole application
requirements.txt        Python dependencies
DEPLOYMENT_GUIDE.md     Keys, GitHub and Streamlit deployment
README.md               This file
```

## Known limitations

- **Free models are slow, rate-limited and may log prompts.** Do not upload confidential or unpublished material to them.
- **Scanned PDFs and image-only slides have no extractable text.** Run OCR first or paste the text. The app does not do OCR.
- **Long input is truncated** to the sidebar's character limit (default 60,000). Raise it only if your model's context window allows.
- **Large exams and decks can be cut off** by the model's output limit. The app keeps the complete items and warns you about the rest.
- **Verify citations.** The analyzer uses only metadata found in the document and marks missing fields as placeholders. Check them against the publisher's page.
- **AI output can be wrong.** Check summaries, formulas and exam answers against your course materials before relying on them.
- **A shared public deployment spends the owner's key.** See the deployment guide for how to limit this.
