# MSU Texas Catalog Agentic RAG Chatbot - Best Demo Version

This is a simple flat Streamlit project for a public stakeholder demo of a student-facing MSU Texas catalog chatbot.

## What this version includes

- Streamlit chatbot UI for Streamlit Community Cloud
- Page and approved subpage crawling
- Catalog change detection using content hashes
- Query Understanding Agent
- Query Rewriting Agent
- Hybrid Retrieval Agent using BM25 + dense semantic retrieval
- Metadata filtering by topic, student type, heading, source URL, and catalog year
- Reranking Agent
- Answer Generation Agent
- Citation Verification and Guardrail Agent
- MCP-style security / guardrail layer
- Report Creation Agent
- RAG evaluation: MRR, relevance, similarity, groundedness, hallucination risk, citation confidence
- Automated stakeholder PowerPoint and 20+ page DOCX report

## Best demo LLM recommendation

For stakeholder demos, use Groq Llama 3.3 70B Versatile through Streamlit Secrets:

```toml
GROQ_API_KEY = "your_groq_api_key_here"
GROQ_MODEL = "llama-3.3-70b-versatile"
```

The app still runs without any API key using a free local grounded fallback, but the hosted LLM gives better student-facing responses.

## Files

- `streamlit_app.py` - public Streamlit chatbot UI
- `main.py` - one-command autonomous pipeline
- `rag_core.py` - crawler, RAG, agents, evaluation, reports
- `mcp_guardrail_server.py` - MCP-style security/guardrail functions
- `requirements.txt` - dependencies
- `outputs/` - generated files after running

## Run locally

```powershell
python -m venv .venv
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python main.py
streamlit run streamlit_app.py
```

## Deploy on Streamlit Community Cloud

1. Upload the extracted project files to a public GitHub repo.
2. In Streamlit Community Cloud, choose the repo.
3. Set the main file path to:

```text
streamlit_app.py
```

4. Add optional secrets for best answer quality:

```toml
GROQ_API_KEY = "your_groq_api_key_here"
GROQ_MODEL = "llama-3.3-70b-versatile"
```

5. Deploy and share the public `.streamlit.app` URL.

## Best demo stack

- Frontend and public hosting: Streamlit Community Cloud
- LLM: Groq Llama 3.3 70B Versatile for best demo; local grounded fallback for free/no-key mode
- RAG: Agentic Hybrid RAG
- Search: BM25 + dense semantic retrieval
- Embedding: TF-IDF + TruncatedSVD contextual dense vectors for free demo
- Reranking: metadata-aware dense + keyword reranker
- Guardrails: MCP-style source allowlisting, PII checks, citation verification, high-stakes escalation
- Reports: python-pptx and python-docx
