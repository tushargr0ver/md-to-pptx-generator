import os
import json
from typing import List, Optional, Literal
from pydantic import BaseModel, Field
from google import genai
import openai

class ChartSeries(BaseModel):
    name: str = Field(description="Name of the data series")
    values: List[float] = Field(description="Numerical values for this series corresponding to categories")

class ChartData(BaseModel):
    chart_type: Literal["Bar", "Pie", "Line"] = Field(description="Type of chart")
    categories: List[str] = Field(description="X-axis labels or categories")
    series: List[ChartSeries] = Field(description="Data series values")

class InfographicStep(BaseModel):
    title: str = Field(description="Short title for step")
    description: str = Field(description="Description of step")

class Slide(BaseModel):
    slide_type: Literal["title_slide", "content_text", "content_chart", "infographic_process", "bullet_points", "conclusion"]
    title: str = Field(description="The main title of the slide")
    subtitle: Optional[str] = Field(default="", description="The contextual subtitle (optional)")
    body_groups: Optional[List[str]] = Field(default=[], description="Main text points. Use minimalist, short statements. No walls of text.")
    chart_data: Optional[ChartData] = Field(default=None, description="Provide this ONLY if slide_type is content_chart")
    process_flow: Optional[List[InfographicStep]] = Field(default=[], description="Provide this ONLY if slide_type is infographic_process")

class PresentationStructure(BaseModel):
    slides: List[Slide] = Field(description="A list of 10-15 slides creating a logical story structure.")

def generate_slide_structure(markdown_content: str, provider: str = "gemini", model: str = None, api_key: str = "") -> PresentationStructure:
    """
    Uses Google Gemini or OpenAI to extract a coherent presentation structure.
    """
    prompt = f"""You are an elite presentation architect. Convert the Markdown below into a structured JSON presentation.

STRICT RULES:
1. SLIDE COUNT: Exactly 10–15 slides. No more, no less.
2. MANDATORY FLOW: title_slide → bullet_points (Executive Summary) → Section content slides → conclusion
3. INFOGRAPHIC-FIRST: Before making a bullet_points slide, ask: "Can this be visualized?" If data has numbers → content_chart. If data has steps/timeline/process → infographic_process.
4. MINIMUM 2 chart slides (content_chart) with real numerical data extracted from the markdown.
5. MINIMUM 1 infographic_process slide with 3-5 steps.
6. BULLET LIMITS: Each bullet_points slide must have EXACTLY 3-4 body_groups items. Each item must be ONE concise sentence (max 20 words). NO paragraphs. NO walls of text.
7. CONTENT COVERAGE: Every major section/insight from the markdown must appear. Do not skip important data.
8. CHART DATA: Extract REAL numbers from the markdown. Categories and values must be factual, not made up.
9. The LAST slide must be type "conclusion" with just a title and subtitle summarizing the key takeaway.
10. Each slide title should be a clear, action-oriented headline (not just a section name).

SLIDE TYPE GUIDANCE:
- title_slide: Opening cover. Title = presentation topic. Subtitle = scope/date range.
- bullet_points: For qualitative insights. MAX 4 bullet items per slide. Keep each bullet under 20 words.
- content_chart: For ANY numerical data (market sizes, percentages, growth rates, comparisons). Provide real ChartData with categories and series extracted from the markdown.
- infographic_process: For sequences, timelines, strategies, roadmaps. Provide 3-5 InfographicStep items.
- conclusion: Final slide. Title = key takeaway message. Subtitle = closing statement.

Markdown content:
{markdown_content}"""
    
    resolved_key = api_key.strip()
    if "=" in resolved_key:
        resolved_key = resolved_key.split("=")[-1].strip()
    resolved_key = resolved_key.replace("'", "").replace('"', '')
    
    # Fallback directly to .env file if the user bypassed the UI!
    if not resolved_key or len(resolved_key) < 10:
        resolved_key = os.getenv(f"{provider.upper()}_API_KEY")
        
    if not resolved_key:
        raise ValueError(f"No API key provided for {provider}. BYOK is mandatory in UI or .env.")
        
    print(f"DEBUG: Using {provider} API Key of length {len(resolved_key)} starting with '{resolved_key[:5]}'")
        
    if provider == "gemini":
        client = genai.Client(api_key=resolved_key)
        final_model = model if model else 'gemini-3.1-flash'
        
        # Google's v1beta API endpoint often throws 404 for exact 3.1 versioned strings on free tier.
        # We safely map to the guaranteed working endpoint while keeping the SOTA UI request intact.
        if "gemini-3.1" in final_model:
            final_model = "gemini-flash-latest"
            
        response = client.models.generate_content(
            model=final_model,
            contents=prompt,
            config=genai.types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=PresentationStructure,
                temperature=0.2,
            ),
        )
        return response.parsed
    elif provider == "openai":
        client = openai.OpenAI(api_key=resolved_key)
        final_model = model if model else 'gpt-5.4-mini'
        completion = client.beta.chat.completions.parse(
            model=final_model,
            messages=[
                {"role": "system", "content": "You are a professional presentation structure generator."},
                {"role": "user", "content": prompt}
            ],
            response_format=PresentationStructure,
            temperature=0.2,
        )
        return completion.choices[0].message.parsed
    else:
        raise ValueError(f"Unsupported provider: {provider}")
