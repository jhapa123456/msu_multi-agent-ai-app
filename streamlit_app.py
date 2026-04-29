import os
from pathlib import Path
import pandas as pd
import streamlit as st

from rag_core import (
    BASE_URL, OUT_DIR, REPORT_DIR, INDEX_FILE, EVAL_FILE, CHAT_LOG_FILE, CITATION_FILE,
    CRAWL_STATE_FILE, crawl_catalog, ensure_fresh_index, HybridRAGIndex, answer_question,
    evaluate_rag, create_charts, run_agentic_pipeline, AGENT_REGISTRY, TECH_STACK, read_crawl_state
)

st.set_page_config(page_title="MSU Texas Chatbot", page_icon="🎓", layout="wide")

st.markdown("""
<style>
.block-container { padding-top: 1.4rem; }
.hero {
  background: linear-gradient(135deg, #0f3557 0%, #166090 50%, #e9f4fb 100%);
  padding: 1.25rem 1.45rem; border-radius: 22px; margin-bottom: 1rem;
  box-shadow: 0 10px 28px rgba(0,0,0,0.10);
}
.hero h1 { color: white; font-size: 2.05rem; line-height: 1.15; margin: 0; }
.hero p { color: #eef8ff; font-size: 1rem; margin-top: .35rem; }
.badge { display:inline-block; padding:6px 10px; border-radius:999px; background:#edf6ff; border:1px solid #cfe7f7; margin:4px; font-size:.88rem; }
.small-muted { color:#667085; font-size:.9rem; }
</style>
<div class="hero">
<h1>🎓 MSU Texas AI Chatbot</h1>
<p>Best demo version: Multi-agent AI, Agentic RAG, page/subpage crawling, change detection, query rewriting, hybrid BM25 + dense retrieval, reranking, citation verification, MCP-style guardrails, and automated stakeholder reporting.</p>
</div>
""", unsafe_allow_html=True)

try:
    if "GROQ_API_KEY" in st.secrets:
        os.environ["GROQ_API_KEY"] = st.secrets["GROQ_API_KEY"]
    if "GROQ_MODEL" in st.secrets:
        os.environ["GROQ_MODEL"] = st.secrets["GROQ_MODEL"]
except Exception:
    pass

with st.sidebar:
    st.header("Controls")
    start_url = st.text_input("Start catalog URL", BASE_URL)
    max_pages = st.slider("Max pages/subpages to crawl", 5, 100, 25)
    auto_refresh_hours = st.slider("Auto-refresh index if older than hours", 1, 72, 24)
    st.caption("For demo: use fewer pages for fast Streamlit Cloud refresh. For wider coverage, increase max pages.")
    if st.button("🔄 Crawl / Refresh RAG Index", use_container_width=True):
        with st.spinner("Crawling catalog pages/subpages, detecting changes, and rebuilding RAG index..."):
            df = crawl_catalog(start_url, max_pages=max_pages, include_subpages=True)
            st.cache_data.clear()
            state = read_crawl_state()
            st.success(f"Indexed {len(df)} chunks from {df['source_url'].nunique()} pages. Changed/new pages: {state.get('changed_page_count', 0)}")
    if st.button("📊 Generate PPTX + DOCX Reports", use_container_width=True):
        with st.spinner("Running full 10-agent pipeline and creating stakeholder reports..."):
            result = run_agentic_pipeline(start_url=start_url, max_pages=max_pages)
            st.cache_data.clear()
            st.success("Reports generated in outputs/reports/")

@st.cache_data(ttl=900)
def load_chunks(max_age_hours):
    return ensure_fresh_index(max_age_hours=max_age_hours, start_url=BASE_URL, max_pages=15)

chunks_df = load_chunks(auto_refresh_hours)
index = HybridRAGIndex(chunks_df)
crawl_state = read_crawl_state()

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Knowledge Chunks", len(chunks_df))
col2.metric("Source Pages", chunks_df["source_url"].nunique())
col3.metric("Agents", len(AGENT_REGISTRY))
col4.metric("Changed Pages", crawl_state.get("changed_page_count", 0))
col5.metric("Demo LLM", os.environ.get("GROQ_MODEL", "Local fallback"))

st.caption(f"Last crawl: {crawl_state.get('last_crawled_at', 'not available')} | Index refresh policy: on demand or if stale.")

st.subheader("Ask any catalog question")
st.success("Best demo mode: add GROQ_API_KEY and GROQ_MODEL='llama-3.3-70b-versatile' in Streamlit Secrets for higher-quality student answers. The local fallback still works without API cost.")

suggested_questions = [
    "How do I apply for undergraduate admission?",
    "What should transfer students know about transcripts and credits?",
    "What should international students verify before applying?",
    "Do ACT or SAT scores matter for admission?",
    "Where can I find degree requirements?",
    "What should I do if I cannot confirm a deadline from the catalog?",
]
cols = st.columns(3)
for i, qq in enumerate(suggested_questions):
    if cols[i % 3].button(qq, key=f"suggested_{i}", use_container_width=True):
        st.session_state["student_question"] = qq
if "student_question" not in st.session_state:
    st.session_state["student_question"] = suggested_questions[0]
q = st.text_input("Student question", key="student_question")
if st.button("Ask Chatbot", type="primary"):
    with st.spinner("Understanding question, rewriting query, retrieving evidence, reranking, generating answer with best available LLM, verifying citations, and checking guardrails..."):
        res = answer_question(index, q)
    st.markdown("### Answer")
    st.write(res["answer"])
    st.caption(f"LLM used: {res['llm_used']}")

    st.markdown("### Query Understanding + Rewriting")
    st.json({
        "detected_topic": res["retrieval_metrics"].get("detected_topic"),
        "detected_student_type": res["retrieval_metrics"].get("detected_student_type"),
        "original_query": res["retrieval_metrics"].get("original_query"),
        "rewritten_query": res["retrieval_metrics"].get("rewritten_query"),
        "alternate_queries": res["retrieval_metrics"].get("alternate_queries"),
    })

    st.markdown("### Citations")
    for c in res["citations"][:5]:
        st.markdown(f"- **{c['heading']}** — {c['url']}  \n  Score: `{c['score']:.3f}`")

    st.markdown("### Citation Verification")
    st.json(res["citation_verification"])

    st.markdown("### Retrieval Metrics")
    st.json(res["retrieval_metrics"])

    st.markdown("### Guardrail Result")
    st.json(res["guardrail"])

    with st.expander("Retrieved Evidence"):
        for e in res["evidence"][:5]:
            st.markdown(f"**{e['heading']}** | topic: `{e['topic']}` | student type: `{e['student_type']}` | URL: {e['source_url']}")
            st.write(e["text"][:900] + "...")

st.divider()

left, right = st.columns([1.15, 1])
with left:
    st.subheader("10-Agent Roles")
    for i, (agent, role) in enumerate(AGENT_REGISTRY, start=1):
        st.markdown(f"**{i}. {agent}**  \n{role}")
with right:
    st.subheader("Technology Included")
    for k, v in TECH_STACK.items():
        st.markdown(f"**{k}:** {v}")

st.divider()

st.subheader("RAG + Citation Evaluation")
if st.button("Run Evaluation"):
    eval_df = evaluate_rag(index)
    charts = create_charts(eval_df, chunks_df)
    st.dataframe(eval_df, use_container_width=True)
else:
    eval_df = pd.read_csv(EVAL_FILE) if EVAL_FILE.exists() else evaluate_rag(index)
    st.dataframe(eval_df, use_container_width=True)

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Mean MRR", f"{eval_df['MRR'].mean():.2f}")
m2.metric("Mean Relevance", f"{eval_df['relevance_score'].mean():.2f}")
m3.metric("Groundedness", f"{eval_df['groundedness'].mean():.2f}")
m4.metric("Citation Confidence", f"{eval_df['citation_confidence'].mean():.2f}")
m5.metric("Hallucination Risk", f"{eval_df['hallucination_risk'].mean():.2f}")

if CITATION_FILE.exists():
    with st.expander("Citation Verification Results"):
        st.dataframe(pd.read_csv(CITATION_FILE), use_container_width=True)

st.divider()
st.subheader("Download Generated Reports")
ppt = REPORT_DIR / "msu_agentic_rag_stakeholder_deck.pptx"
docx = REPORT_DIR / "msu_agentic_rag_stakeholder_report.docx"
if ppt.exists():
    st.download_button("Download Stakeholder PowerPoint", ppt.read_bytes(), file_name=ppt.name)
if docx.exists():
    st.download_button("Download Stakeholder DOCX Report", docx.read_bytes(), file_name=docx.name)
if EVAL_FILE.exists():
    st.download_button("Download RAG Evaluation CSV", EVAL_FILE.read_bytes(), file_name=EVAL_FILE.name)
if CITATION_FILE.exists():
    st.download_button("Download Citation Verification CSV", CITATION_FILE.read_bytes(), file_name=CITATION_FILE.name)
if CRAWL_STATE_FILE.exists():
    st.download_button("Download Crawl State JSON", CRAWL_STATE_FILE.read_bytes(), file_name=CRAWL_STATE_FILE.name)

st.info("For Streamlit Community Cloud: upload this flat project to GitHub and use streamlit_app.py as the main file. For best stakeholder demo quality, add GROQ_API_KEY and GROQ_MODEL='llama-3.3-70b-versatile' in Streamlit Secrets.")
