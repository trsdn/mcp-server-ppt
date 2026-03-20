"""
Batch slide generator for evaluation.

For each prompt, generates a 2-slide presentation (title + key content slide)
using the current design skills as guidance, exports as PNG.

Run: python eval/generate_batch.py [start_index] [count]
Default: generates all 100 prompts
"""

import json
import subprocess
import os
import sys
import time
from pathlib import Path

REPO_ROOT = Path(__file__).parent.parent
EVAL_ASSET_REPO_ROOT_ENVIRONMENT_VARIABLE = "PPTMCP_EVAL_ASSET_REPO_ROOT"
PPTCLI = str(REPO_ROOT / "src" / "PptMcp.CLI" / "bin" / "Release" / "net9.0-windows" / "pptcli.exe")


def get_eval_asset_repo_root() -> Path:
    configured_root = os.environ.get(EVAL_ASSET_REPO_ROOT_ENVIRONMENT_VARIABLE)
    return Path(configured_root).resolve() if configured_root else REPO_ROOT


EVAL_ASSET_REPO_ROOT = get_eval_asset_repo_root()
OUTPUT_DIR = EVAL_ASSET_REPO_ROOT / "eval" / "output"
PROMPTS_FILE = REPO_ROOT / "eval" / "prompts" / "test-prompts.json"

# Color profiles by category
PROFILES = {
    "corporate": {"bg": "#0B3D91", "accent": "#FF6B35", "text": "#1A1A2E", "primary": "#0B3D91", "positive": "#2E8B57", "neutral": "#F0F4F8", "font": "Calibri"},
    "sales": {"bg": "#2C3E50", "accent": "#E74C3C", "text": "#2C3E50", "primary": "#2C3E50", "positive": "#27AE60", "neutral": "#ECF0F1", "font": "Segoe UI"},
    "project": {"bg": "#2D3748", "accent": "#3182CE", "text": "#1A202C", "primary": "#2D3748", "positive": "#38A169", "neutral": "#EDF2F7", "font": "Calibri"},
    "executive": {"bg": "#0B3D91", "accent": "#FF6B35", "text": "#1A1A2E", "primary": "#0B3D91", "positive": "#2E8B57", "neutral": "#F0F4F8", "font": "Calibri"},
    "product": {"bg": "#6C5CE7", "accent": "#FDCB6E", "text": "#2D3436", "primary": "#6C5CE7", "positive": "#00B894", "neutral": "#F0F0F5", "font": "Segoe UI"},
    "training": {"bg": "#0B3D91", "accent": "#FF6B35", "text": "#1A1A2E", "primary": "#0B3D91", "positive": "#2E8B57", "neutral": "#F0F4F8", "font": "Calibri"},
    "dashboard": {"bg": "#0B3D91", "accent": "#FF6B35", "text": "#1A1A2E", "primary": "#0B3D91", "positive": "#2E8B57", "neutral": "#F0F4F8", "font": "Calibri"},
    "strategy": {"bg": "#0B3D91", "accent": "#FF6B35", "text": "#1A1A2E", "primary": "#0B3D91", "positive": "#2E8B57", "neutral": "#F0F4F8", "font": "Calibri"},
    "team": {"bg": "#2D3748", "accent": "#3182CE", "text": "#1A202C", "primary": "#2D3748", "positive": "#38A169", "neutral": "#EDF2F7", "font": "Calibri"},
    "creative": {"bg": "#2C3E50", "accent": "#FF6B6B", "text": "#2C3E50", "primary": "#FF6B6B", "positive": "#4ECDC4", "neutral": "#F7F7F7", "font": "Segoe UI"},
    "technical": {"bg": "#2D3748", "accent": "#3182CE", "text": "#1A202C", "primary": "#2D3748", "positive": "#38A169", "neutral": "#EDF2F7", "font": "Calibri"},
}


def cli(*args):
    """Run pptcli command, return parsed JSON."""
    result = subprocess.run([PPTCLI] + list(args), capture_output=True, text=True, timeout=30)
    # Find last JSON line
    for line in reversed(result.stdout.strip().split("\n")):
        line = line.strip()
        if line.startswith("{"):
            return json.loads(line)
    return {"success": False, "error": result.stderr or result.stdout}


def generate_title_slide(session_id: str, prompt: dict, profile: dict):
    """Generate a dark hero title slide."""
    title = prompt["prompt"][:60]  # Truncate for title
    if ":" in title:
        title = title.split(":")[0].strip()
    
    s = session_id
    cli("slide", "create", "-s", s, "--layout-name", "Blank", "--position", "1")
    
    # Dark background
    cli("shape", "add-shape", "-s", s, "--slide-index", "1", "--auto-shape-type", "1",
        "--left", "0", "--top", "0", "--width", "960", "--height", "540")
    cli("shape", "set-fill", "-s", s, "--slide-index", "1", "--shape-name", "Rectangle 1", "--color-hex", profile["bg"])
    cli("shape", "set-line", "-s", s, "--slide-index", "1", "--shape-name", "Rectangle 1", "--color-hex", "none")
    
    # Accent bar
    cli("shape", "add-shape", "-s", s, "--slide-index", "1", "--auto-shape-type", "1",
        "--left", "60", "--top", "148", "--width", "80", "--height", "4")
    cli("shape", "set-fill", "-s", s, "--slide-index", "1", "--shape-name", "Rectangle 2", "--color-hex", profile["accent"])
    cli("shape", "set-line", "-s", s, "--slide-index", "1", "--shape-name", "Rectangle 2", "--color-hex", "none")
    
    # Title
    cli("shape", "add-textbox", "-s", s, "--slide-index", "1",
        "--left", "60", "--top", "160", "--width", "840", "--height", "70", "--text", title)
    cli("text", "format", "-s", s, "--slide-index", "1", "--shape-name", "TextBox 3",
        "--font-size", "36", "--bold", "true", "--color", "#FFFFFF", "--font-name", profile["font"])
    
    # Subtitle
    subtitle = prompt["prompt"][60:140] if len(prompt["prompt"]) > 60 else prompt["category"].title() + " Presentation"
    cli("shape", "add-textbox", "-s", s, "--slide-index", "1",
        "--left", "60", "--top", "240", "--width", "600", "--height", "30", "--text", subtitle)
    cli("text", "format", "-s", s, "--slide-index", "1", "--shape-name", "TextBox 4",
        "--font-size", "18", "--color", "#B0C4DE", "--font-name", profile["font"])
    
    # Date + category
    cli("shape", "add-textbox", "-s", s, "--slide-index", "1",
        "--left", "60", "--top", "460", "--width", "300", "--height", "20", "--text", f"March 2026 | {prompt['category'].title()}")
    cli("text", "format", "-s", s, "--slide-index", "1", "--shape-name", "TextBox 5",
        "--font-size", "12", "--color", "#7B9CC0", "--font-name", profile["font"])
    
    # Bottom accent bar
    cli("shape", "add-shape", "-s", s, "--slide-index", "1", "--auto-shape-type", "1",
        "--left", "0", "--top", "520", "--width", "960", "--height", "20")
    cli("shape", "set-fill", "-s", s, "--slide-index", "1", "--shape-name", "Rectangle 6", "--color-hex", profile["accent"])
    cli("shape", "set-line", "-s", s, "--slide-index", "1", "--shape-name", "Rectangle 6", "--color-hex", "none")


def generate_content_slide(session_id: str, prompt: dict, profile: dict):
    """Generate a content slide appropriate for the category."""
    s = session_id
    cli("slide", "create", "-s", s, "--layout-name", "Blank", "--position", "2")
    
    # Action title
    action_title = _derive_action_title(prompt)
    cli("shape", "add-textbox", "-s", s, "--slide-index", "2",
        "--left", "36", "--top", "20", "--width", "888", "--height", "45", "--text", action_title)
    cli("text", "format", "-s", s, "--slide-index", "2", "--shape-name", "TextBox 1",
        "--font-size", "18", "--bold", "true", "--color", profile["text"], "--font-name", profile["font"])
    
    # Accent underline
    cli("shape", "add-shape", "-s", s, "--slide-index", "2", "--auto-shape-type", "1",
        "--left", "36", "--top", "68", "--width", "80", "--height", "3")
    cli("shape", "set-fill", "-s", s, "--slide-index", "2", "--shape-name", "Rectangle 2", "--color-hex", profile["accent"])
    cli("shape", "set-line", "-s", s, "--slide-index", "2", "--shape-name", "Rectangle 2", "--color-hex", "none")
    
    # Content based on category
    cat = prompt["category"]
    if cat in ("dashboard", "corporate"):
        _generate_kpi_cards(s, prompt, profile)
    elif cat in ("strategy", "executive"):
        _generate_pillars(s, prompt, profile)
    elif cat in ("sales", "product"):
        _generate_features(s, prompt, profile)
    elif cat in ("project",):
        _generate_status_cards(s, prompt, profile)
    else:
        _generate_bullet_list(s, prompt, profile)
    
    # Page number
    cli("shape", "add-textbox", "-s", s, "--slide-index", "2",
        "--left", "900", "--top", "515", "--width", "40", "--height", "18", "--text", "2")
    cli("text", "format", "-s", s, "--slide-index", "2", "--shape-name", _last_textbox(s, 2),
        "--font-size", "9", "--color", "#999999", "--alignment", "right")


def _derive_action_title(prompt: dict) -> str:
    """Create an action title from the prompt."""
    text = prompt["prompt"]
    # Extract key numbers/claims if present
    if "$" in text or "%" in text:
        return text[:80]
    return text[:80]


def _last_textbox(session_id, slide_index):
    """Get the name of the last textbox added."""
    result = cli("shape", "list", "-s", session_id, "--slide-index", str(slide_index))
    if result.get("success") and result.get("shapes"):
        shapes = result["shapes"]
        textboxes = [s for s in shapes if "TextBox" in s.get("name", "")]
        if textboxes:
            return textboxes[-1]["name"]
    return "TextBox 99"


def _generate_kpi_cards(session_id: str, prompt: dict, profile: dict):
    """Generate 3 KPI cards across top."""
    s = session_id
    cards = [
        {"label": "KEY METRIC 1", "value": "99.7%", "context": "▲ On Track", "color": profile["primary"]},
        {"label": "KEY METRIC 2", "value": "▼ 23%", "context": "vs. Last Period", "color": profile["positive"]},
        {"label": "KEY METRIC 3", "value": "4.6/5", "context": "Target: 4.5", "color": profile["accent"]},
    ]
    x_positions = [36, 332, 628]
    
    for i, (card, x) in enumerate(zip(cards, x_positions)):
        # Card background
        bg_name = f"Rectangle: Rounded Corners {3 + i*4}"
        cli("shape", "add-shape", "-s", s, "--slide-index", "2", "--auto-shape-type", "5",
            "--left", str(x), "--top", "85", "--width", "276", "--height", "180")
        cli("shape", "set-fill", "-s", s, "--slide-index", "2", "--shape-name", bg_name, "--color-hex", profile["neutral"])
        cli("shape", "set-line", "-s", s, "--slide-index", "2", "--shape-name", bg_name, "--color-hex", "none")
        
        # Header strip
        rect_name = f"Rectangle {4 + i*4}"
        cli("shape", "add-shape", "-s", s, "--slide-index", "2", "--auto-shape-type", "1",
            "--left", str(x), "--top", "85", "--width", "276", "--height", "28")
        cli("shape", "set-fill", "-s", s, "--slide-index", "2", "--shape-name", rect_name, "--color-hex", card["color"])
        cli("shape", "set-line", "-s", s, "--slide-index", "2", "--shape-name", rect_name, "--color-hex", "none")
        cli("text", "set", "-s", s, "--slide-index", "2", "--shape-name", rect_name, "--text", card["label"])
        cli("text", "format", "-s", s, "--slide-index", "2", "--shape-name", rect_name,
            "--font-size", "10", "--bold", "true", "--color", "#FFFFFF")
        
        # Big number
        tb_val = f"TextBox {5 + i*4}"
        cli("shape", "add-textbox", "-s", s, "--slide-index", "2",
            "--left", str(x), "--top", "125", "--width", "276", "--height", "55", "--text", card["value"])
        cli("text", "format", "-s", s, "--slide-index", "2", "--shape-name", tb_val,
            "--font-size", "44", "--bold", "true", "--color", card["color"], "--alignment", "center")
        
        # Context
        tb_ctx = f"TextBox {6 + i*4}"
        cli("shape", "add-textbox", "-s", s, "--slide-index", "2",
            "--left", str(x), "--top", "190", "--width", "276", "--height", "25", "--text", card["context"])
        cli("text", "format", "-s", s, "--slide-index", "2", "--shape-name", tb_ctx,
            "--font-size", "12", "--color", "#4A5568", "--alignment", "center")


def _generate_pillars(session_id: str, prompt: dict, profile: dict):
    """Generate 3-column strategic pillars."""
    s = session_id
    pillars = [
        {"title": "Pillar 1", "color": profile["primary"], "items": "• Initiative A\n• Initiative B\n• Initiative C"},
        {"title": "Pillar 2", "color": profile["accent"], "items": "• Initiative D\n• Initiative E\n• Initiative F"},
        {"title": "Pillar 3", "color": profile["positive"], "items": "• Initiative G\n• Initiative H\n• Initiative I"},
    ]
    x_positions = [36, 332, 628]
    
    for i, (pillar, x) in enumerate(zip(pillars, x_positions)):
        # Card
        bg_name = f"Rectangle: Rounded Corners {3 + i*5}"
        cli("shape", "add-shape", "-s", s, "--slide-index", "2", "--auto-shape-type", "5",
            "--left", str(x), "--top", "85", "--width", "276", "--height", "380")
        cli("shape", "set-fill", "-s", s, "--slide-index", "2", "--shape-name", bg_name, "--color-hex", profile["neutral"])
        cli("shape", "set-line", "-s", s, "--slide-index", "2", "--shape-name", bg_name, "--color-hex", "none")
        
        # Number badge
        oval_name = f"Oval {4 + i*5}"
        cli("shape", "add-shape", "-s", s, "--slide-index", "2", "--auto-shape-type", "9",
            "--left", str(x + 113), "--top", "100", "--width", "50", "--height", "50")
        cli("shape", "set-fill", "-s", s, "--slide-index", "2", "--shape-name", oval_name, "--color-hex", pillar["color"])
        cli("shape", "set-line", "-s", s, "--slide-index", "2", "--shape-name", oval_name, "--color-hex", "none")
        cli("text", "set", "-s", s, "--slide-index", "2", "--shape-name", oval_name, "--text", str(i + 1))
        cli("text", "format", "-s", s, "--slide-index", "2", "--shape-name", oval_name,
            "--font-size", "20", "--bold", "true", "--color", "#FFFFFF", "--alignment", "center")
        
        # Title
        tb_title = f"TextBox {5 + i*5}"
        cli("shape", "add-textbox", "-s", s, "--slide-index", "2",
            "--left", str(x + 20), "--top", "160", "--width", "236", "--height", "30", "--text", pillar["title"])
        cli("text", "format", "-s", s, "--slide-index", "2", "--shape-name", tb_title,
            "--font-size", "16", "--bold", "true", "--color", pillar["color"], "--alignment", "center")
        
        # Bullet items
        tb_items = f"TextBox {6 + i*5}"
        cli("shape", "add-textbox", "-s", s, "--slide-index", "2",
            "--left", str(x + 20), "--top", "200", "--width", "236", "--height", "230", "--text", pillar["items"])
        cli("text", "format", "-s", s, "--slide-index", "2", "--shape-name", tb_items,
            "--font-size", "12", "--color", "#333333")


def _generate_features(session_id: str, prompt: dict, profile: dict):
    """Generate 3-column feature cards for sales/product."""
    s = session_id
    features = [
        {"icon": "A", "title": "Feature 1", "desc": "Key benefit and value proposition for the customer."},
        {"icon": "B", "title": "Feature 2", "desc": "Another important capability that differentiates us."},
        {"icon": "C", "title": "Feature 3", "desc": "Third major advantage driving customer adoption."},
    ]
    x_positions = [36, 332, 628]
    
    for i, (feat, x) in enumerate(zip(features, x_positions)):
        # Card
        bg_name = f"Rectangle: Rounded Corners {3 + i*5}"
        cli("shape", "add-shape", "-s", s, "--slide-index", "2", "--auto-shape-type", "5",
            "--left", str(x), "--top", "85", "--width", "276", "--height", "350")
        cli("shape", "set-fill", "-s", s, "--slide-index", "2", "--shape-name", bg_name, "--color-hex", profile["neutral"])
        cli("shape", "set-line", "-s", s, "--slide-index", "2", "--shape-name", bg_name, "--color-hex", "none")
        
        # Icon circle
        oval_name = f"Oval {4 + i*5}"
        cli("shape", "add-shape", "-s", s, "--slide-index", "2", "--auto-shape-type", "9",
            "--left", str(x + 113), "--top", "110", "--width", "50", "--height", "50")
        colors = [profile["accent"], profile["positive"], profile["primary"]]
        cli("shape", "set-fill", "-s", s, "--slide-index", "2", "--shape-name", oval_name, "--color-hex", colors[i])
        cli("shape", "set-line", "-s", s, "--slide-index", "2", "--shape-name", oval_name, "--color-hex", "none")
        cli("text", "set", "-s", s, "--slide-index", "2", "--shape-name", oval_name, "--text", feat["icon"])
        cli("text", "format", "-s", s, "--slide-index", "2", "--shape-name", oval_name,
            "--font-size", "18", "--bold", "true", "--color", "#FFFFFF", "--alignment", "center")
        
        # Title
        tb_title = f"TextBox {5 + i*5}"
        cli("shape", "add-textbox", "-s", s, "--slide-index", "2",
            "--left", str(x + 20), "--top", "180", "--width", "236", "--height", "30", "--text", feat["title"])
        cli("text", "format", "-s", s, "--slide-index", "2", "--shape-name", tb_title,
            "--font-size", "16", "--bold", "true", "--color", profile["text"], "--alignment", "center")
        
        # Description
        tb_desc = f"TextBox {6 + i*5}"
        cli("shape", "add-textbox", "-s", s, "--slide-index", "2",
            "--left", str(x + 20), "--top", "220", "--width", "236", "--height", "150", "--text", feat["desc"])
        cli("text", "format", "-s", s, "--slide-index", "2", "--shape-name", tb_desc,
            "--font-size", "13", "--color", "#666666", "--alignment", "center")


def _generate_status_cards(session_id: str, prompt: dict, profile: dict):
    """Generate status/RAG cards for project prompts."""
    _generate_kpi_cards(session_id, prompt, profile)


def _generate_bullet_list(session_id: str, prompt: dict, profile: dict):
    """Generate a clean bullet list slide."""
    s = session_id
    bullets = "• Key point one with supporting detail\n\n• Key point two with context and evidence\n\n• Key point three with actionable recommendation\n\n• Key point four with measurable outcome"
    
    cli("shape", "add-textbox", "-s", s, "--slide-index", "2",
        "--left", "60", "--top", "90", "--width", "840", "--height", "380", "--text", bullets)
    cli("text", "format", "-s", s, "--slide-index", "2", "--shape-name", "TextBox 3",
        "--font-size", "16", "--color", profile["text"], "--font-name", profile["font"])


def generate_presentation(prompt: dict) -> str:
    """Generate a 2-slide presentation for a prompt. Returns output PNG path."""
    pid = prompt["id"]
    category = prompt["category"]
    profile = PROFILES.get(category, PROFILES["corporate"])
    
    pptx_path = str(OUTPUT_DIR / f"{pid}.pptx")
    png_path = str(OUTPUT_DIR / f"{pid}.png")
    
    # Clean up existing
    for p in [pptx_path, png_path]:
        if os.path.exists(p):
            os.remove(p)
    
    # Create presentation
    result = cli("session", "create", pptx_path)
    if not result.get("success"):
        print(f"  ERROR creating {pid}: {result.get('error', 'unknown')}")
        return ""
    
    session_id = result["sessionId"]
    
    try:
        generate_title_slide(session_id, prompt, profile)
        generate_content_slide(session_id, prompt, profile)
        
        # Export slide 2 (the content slide - more interesting than title)
        cli("export", "slide-to-image", "-s", session_id, "--slide-index", "2", "--destination-path", png_path)
        
        cli("session", "close", "-s", session_id, "--save")
    except Exception as e:
        print(f"  ERROR building {pid}: {e}")
        try:
            cli("session", "close", "-s", session_id)
        except:
            pass
        return ""
    
    return png_path


def main():
    start = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    count = int(sys.argv[2]) if len(sys.argv) > 2 else 100
    
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    
    # Ensure service is running
    cli("service", "start")
    time.sleep(2)
    
    with open(PROMPTS_FILE, "r") as f:
        prompts = json.load(f)
    
    prompts = prompts[start:start + count]
    
    print(f"Generating {len(prompts)} presentations (start={start})")
    
    generated = 0
    failed = 0
    
    for i, prompt in enumerate(prompts):
        pid = prompt["id"]
        print(f"[{i+1}/{len(prompts)}] {pid} ({prompt['category']})...", end=" ", flush=True)
        
        png = generate_presentation(prompt)
        if png and os.path.exists(png):
            size_kb = os.path.getsize(png) / 1024
            print(f"OK ({size_kb:.0f}KB)")
            generated += 1
        else:
            print("FAILED")
            failed += 1
        
        # Small delay to let COM settle
        time.sleep(0.5)
    
    print(f"\nDone: {generated} generated, {failed} failed")


if __name__ == "__main__":
    main()
