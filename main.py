from rag_core import run_agentic_pipeline, BASE_URL

if __name__ == "__main__":
    print("Starting MSU Texas Dynamic Catalog Agentic RAG pipeline...")
    result = run_agentic_pipeline(start_url=BASE_URL, max_pages=35)
    summary = result["summary"]
    print("\nDONE")
    print(f"Chunks indexed: {summary['chunks']}")
    print(f"Source pages crawled: {summary['source_pages']}")
    print(f"Agents: {summary['agents']}")
    print(f"Mean MRR: {summary['mean_MRR']:.2f}")
    print(f"Mean groundedness: {summary['mean_groundedness']:.2f}")
    print("\nGenerated outputs:")
    print("outputs/catalog_chunks.csv")
    print("outputs/rag_evaluation_results.csv")
    print("outputs/agent_chat_log.json")
    print("outputs/charts/")
    print("outputs/reports/msu_agentic_rag_stakeholder_deck.pptx")
    print("outputs/reports/msu_agentic_rag_stakeholder_report.docx")
