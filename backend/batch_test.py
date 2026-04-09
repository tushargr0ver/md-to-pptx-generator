"""
Batch processor for all 24 test cases.
Generates PPTX files with proper rate-limit handling.
Output structure: generated_outputs/<name>/<name>.md + <name>.pptx
"""

import os
import sys
import time
import shutil
import traceback
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

from pipeline.StorytellerAgent import generate_slide_structure
from pipeline.PPTXRenderer import PPTXRenderer

# Paths
TEST_CASES_DIR = Path("../notes/Code EZ_ Master of Agents _ Files-20260409T100059Z-3-001/Code EZ_ Master of Agents _ Files/Test Cases")
MASTER_PATH = Path(__file__).parent / "assets" / "Template.pptx"
OUTPUT_DIR = Path("../generated_outputs")

# Limits
MAX_CHARS = 400_000  # ~100K tokens, well within 1M TPM limit for Flash
COOLDOWN_SECONDS = 45  # Increased to 45s for better stability
MAX_RETRIES = 5        # Increased to 5 retries

def process_file(md_path: Path, output_dir: Path):
    """Process a single markdown file into a PPTX."""
    name = md_path.stem
    
    # Create output folder
    folder = output_dir / name
    pptx_path = folder / f"{name}.pptx"
    
    # Skip if already exists
    if pptx_path.exists():
        print(f"  ⏭ Skipping: {name}.pptx already exists")
        return True
    
    # Read and optionally truncate
    md_text = md_path.read_text(encoding="utf-8")
    original_len = len(md_text)
    
    if len(md_text) > MAX_CHARS:
        # Smart truncation: keep beginning (context) + end (conclusions)
        head_size = int(MAX_CHARS * 0.7)  # 70% from beginning 
        tail_size = int(MAX_CHARS * 0.3)  # 30% from end
        md_text = md_text[:head_size] + "\n\n[... content condensed for processing ...]\n\n" + md_text[-tail_size:]
        print(f"  ⚠ Truncated: {original_len:,} → {len(md_text):,} chars")
    
    # Create output folder
    folder = output_dir / name
    folder.mkdir(parents=True, exist_ok=True)
    
    # Copy original MD file
    shutil.copy2(md_path, folder / md_path.name)
    
    # Call Gemini with retries
    slides_data = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            print(f"  Attempt {attempt}/{MAX_RETRIES}...")
            slides_data = generate_slide_structure(md_text)
            break
        except Exception as e:
            err_msg = str(e)
            if "503" in err_msg or "UNAVAILABLE" in err_msg:
                wait = COOLDOWN_SECONDS * attempt
                print(f"  ⏳ 503 error, waiting {wait}s before retry...")
                time.sleep(wait)
            elif "429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg:
                wait = 60 * attempt
                print(f"  ⏳ Rate limited, waiting {wait}s before retry...")
                time.sleep(wait)
            else:
                print(f"  ✗ Error: {err_msg[:100]}")
                return False
    
    if slides_data is None:
        print(f"  ✗ Failed after {MAX_RETRIES} retries")
        return False
    
    # Render PPTX
    print(f"  Rendering {len(slides_data.slides)} slides...")
    renderer = PPTXRenderer(MASTER_PATH)
    output = renderer.render_slides(slides_data)
    
    pptx_path = folder / f"{name}.pptx"
    with open(pptx_path, "wb") as f:
        f.write(output.getvalue())
    
    print(f"  ✓ Saved: {pptx_path.name} ({len(output.getvalue()) / 1024:.0f} KB)")
    return True


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    
    # Get all .md files sorted by size (smallest first = fastest)
    md_files = sorted(TEST_CASES_DIR.glob("*.md"), key=lambda f: f.stat().st_size)
    
    print(f"{'='*60}")
    print(f"BATCH PROCESSING: {len(md_files)} test cases")
    print(f"Output: {OUTPUT_DIR.resolve()}")
    print(f"Cooldown: {COOLDOWN_SECONDS}s between files")
    print(f"{'='*60}\n")
    
    results = {"success": [], "failed": []}
    
    for i, md_file in enumerate(md_files):
        size_kb = md_file.stat().st_size / 1024
        print(f"\n[{i+1}/{len(md_files)}] {md_file.name} ({size_kb:.0f} KB)")
        
        try:
            success = process_file(md_file, OUTPUT_DIR)
            if success:
                results["success"].append(md_file.name)
            else:
                results["failed"].append(md_file.name)
        except Exception as e:
            print(f"  ✗ Unexpected error: {e}")
            traceback.print_exc()
            results["failed"].append(md_file.name)
        
        # Cooldown between files (skip after last file)
        if i < len(md_files) - 1:
            print(f"  ⏳ Cooling down {COOLDOWN_SECONDS}s...")
            time.sleep(COOLDOWN_SECONDS)
    
    # Summary
    print(f"\n{'='*60}")
    print(f"BATCH COMPLETE")
    print(f"  ✓ Success: {len(results['success'])}/{len(md_files)}")
    print(f"  ✗ Failed:  {len(results['failed'])}/{len(md_files)}")
    if results["failed"]:
        print(f"\nFailed files:")
        for f in results["failed"]:
            print(f"  - {f}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
