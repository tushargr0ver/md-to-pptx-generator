import os
import io
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from dotenv import load_dotenv

from pipeline.MarkdownParser import parse_markdown
from pipeline.StorytellerAgent import generate_slide_structure
from pipeline.LayoutManager import LayoutManager
from pipeline.PPTXRenderer import PPTXRenderer

load_dotenv()

app = FastAPI(title="MD to PPTX AI Agent")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/api/generate")
async def generate_pptx(
    markdown_file: UploadFile = File(...),
    provider: str = Form("gemini"),
    model: str = Form(""),
    api_key: str = Form("")
):
    try:
        # 1. Parse Markdown
        if not markdown_file.filename.endswith(".md"):
            raise HTTPException(status_code=400, detail="Must be a Markdown (.md) file.")
        
        md_content = await markdown_file.read()
        md_text = md_content.decode("utf-8")
        
        print(f"Agent thinking via {provider}...")
        slides_data = generate_slide_structure(
            md_text, 
            provider=provider, 
            model=model, 
            api_key=api_key
        )
        
        # 3. Initialize Presentation renderer
        print("Rendering PPTX...")
        # Since we should use the Slide Master, we'll assume there's a default one in the notes folder for now,
        # but in a complete flow we might allow uploading it.
        # For simplicity in this endpoint:
        default_master = os.path.join(os.path.dirname(__file__), "assets", "Template.pptx")
        
        renderer = PPTXRenderer(default_master)
        presentation_io = renderer.render_slides(slides_data)
        
        # 4. Return as downloadable PPTX
        return StreamingResponse(
            presentation_io,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": f"attachment; filename=generated_{markdown_file.filename}.pptx"}
        )
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
