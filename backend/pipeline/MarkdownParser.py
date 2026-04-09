def parse_markdown(md_text: str) -> str:
    # Basic cleanup in case we need to filter out metadata or extreme lengths.
    # The hackathon brief requires supporting files up to 5MB, which easily fits in Gemini's 1M-2M context window.
    # We strip extra spaces or simple noise if needed, but for now we return it directly so the LLM has all context.
    return md_text
