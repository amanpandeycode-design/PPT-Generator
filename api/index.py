"""
AI PPT Generator - Backend
---------------------------
Flask server that:
  1. Takes a topic + theme + slide count from the frontend
  2. Calls Claude to generate a structured slide outline (title, bullets,
     whether an image is needed, whether a flowchart is needed + its steps)
  3. Builds a real .pptx file using python-pptx (themed colors/fonts,
     auto-drawn flowcharts as native shapes, generated placeholder art)
  4. Returns the .pptx file as a download

Setup:
    pip install -r requirements.txt
    export ANTHROPIC_API_KEY=sk-ant-...
    python main.py
    -> server runs on http://localhost:5000
"""

import os
import io
import json
import random
import urllib.parse
import requests
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from google import genai
from google.genai import types

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from PIL import Image, ImageDraw, ImageFilter
import textwrap

app = Flask(__name__)
CORS(app)

MODEL = "gemini-2.5-flash"  # free tier, no credit card required — https://aistudio.google.com/apikey
_client = None


def get_client():
    """Lazily create the Gemini client so the server can still start (and
    serve routes like /api/themes or /api/generate-pptx) even before a
    GEMINI_API_KEY is set. The key is only required when outline
    generation is actually called."""
    global _client
    if _client is None:
        api_key = os.environ.get("GEMINI_API_KEY")
        if not api_key:
            raise RuntimeError(
                "GEMINI_API_KEY is not set. Get a free key at "
                "https://aistudio.google.com/apikey and export it before generating an outline."
            )
        _client = genai.Client(api_key=api_key)
    return _client

# ---------------------------------------------------------------------------
# THEMES - each theme drives colors, fonts and placeholder-art gradients
# ---------------------------------------------------------------------------
THEMES = {
    "midnight": {
        "name": "Midnight",
        "bg": "0F1220", "primary": "6C5CE7", "secondary": "00CEC9",
        "text": "F5F6FA", "muted": "A0A3BD", "font": "Poppins",
        "gradient": ["1A1C2E", "2D2F4A"],
    },
    "sunrise": {
        "name": "Sunrise",
        "bg": "FFF8F0", "primary": "FF6B35", "secondary": "F7C548",
        "text": "2B2118", "muted": "8A7968", "font": "Poppins",
        "gradient": ["FFDDB3", "FFB570"],
    },
    "forest": {
        "name": "Forest",
        "bg": "F4F9F4", "primary": "1B6B3E", "secondary": "8FBF6E",
        "text": "18291D", "muted": "5F7A65", "font": "Poppins",
        "gradient": ["C9E4CA", "8FBF6E"],
    },
    "ocean": {
        "name": "Ocean",
        "bg": "F0F7FF", "primary": "0B6E99", "secondary": "37B6E8",
        "text": "0B1F2A", "muted": "5C7A8A", "font": "Poppins",
        "gradient": ["BFE3F5", "6FC3E8"],
    },
    "mono": {
        "name": "Monochrome",
        "bg": "FFFFFF", "primary": "1A1A1A", "secondary": "6E6E6E",
        "text": "111111", "muted": "8C8C8C", "font": "Poppins",
        "gradient": ["E8E8E8", "CFCFCF"],
    },
}


def hexc(h):
    return RGBColor.from_string(h)


# ---------------------------------------------------------------------------
# REAL AI IMAGE GENERATION (Pollinations.ai — free, no API key required)
# Returns a direct image URL that both the frontend preview AND the final
# pptx builder use, so what you see in preview is exactly what you get.
# ---------------------------------------------------------------------------
def image_url_for(query, theme, w=1000, h=1000):
    prompt = f"{query}, {theme['name']} color palette, professional presentation graphic, high quality, no text, no watermark"
    encoded = urllib.parse.quote(prompt)
    seed = abs(hash(query)) % 100000
    return f"https://image.pollinations.ai/prompt/{encoded}?width={w}&height={h}&seed={seed}&nologo=true"


def fetch_image_bytes(url, timeout=25):
    """Download the real image for embedding in the pptx. Falls back to
    generated placeholder art if the request fails (offline, timeout, etc.)."""
    resp = requests.get(url, timeout=timeout)
    resp.raise_for_status()
    return io.BytesIO(resp.content)


# ---------------------------------------------------------------------------
# STEP 1: Ask Claude to build the outline
# ---------------------------------------------------------------------------
OUTLINE_SYSTEM_PROMPT = """You are a presentation content strategist. Given a topic, \
produce a slide-by-slide outline as STRICT JSON only (no markdown fences, no preamble).

Schema:
{
  "title": "Deck title",
  "subtitle": "One-line subtitle",
  "slides": [
    {
      "type": "title" | "content" | "image" | "flowchart" | "closing",
      "title": "Slide title",
      "bullets": ["short bullet", "..."],   // 2-5 concise bullets, [] for title/closing
      "needs_image": true|false,             // true if a supporting visual helps
      "image_query": "short visual description",  // required if needs_image
      "needs_flowchart": true|false,         // true if this content is a process/sequence/pipeline
      "flowchart_steps": ["Step 1", "Step 2", "..."]  // 3-6 short steps, required if needs_flowchart
    }
  ]
}

Rules:
- 6 to 10 slides total including a title slide and a closing slide.
- Only mark needs_flowchart true for genuinely process-like content (steps, pipelines, workflows, lifecycles).
- Only mark needs_image true when a concrete visual concept exists, not for every slide.
- Bullets must be short (max ~10 words each), no sub-bullets.
- Return ONLY valid JSON, nothing else."""


def generate_outline(topic, num_slides, audience):
    user_prompt = (
        f"Topic: {topic}\n"
        f"Target slide count: approximately {num_slides}\n"
        f"Audience: {audience or 'general audience'}\n"
        f"Generate the outline JSON now."
    )
    resp = get_client().models.generate_content(
        model=MODEL,
        contents=user_prompt,
        config=types.GenerateContentConfig(
            system_instruction=OUTLINE_SYSTEM_PROMPT,
            response_mime_type="application/json",
            max_output_tokens=3000,
            temperature=0.8,
        ),
    )
    raw = resp.text.strip()
    raw = raw.replace("```json", "").replace("```", "").strip()
    outline = json.loads(raw)

    return outline


def attach_image_urls(outline, theme_key):
    """Post-process the outline: every slide flagged needs_image gets a real
    generated image URL, built from its image_query + the chosen theme."""
    theme = THEMES.get(theme_key, THEMES["midnight"])
    for sl in outline.get("slides", []):
        if sl.get("needs_image") and sl.get("image_query"):
            sl["image_url"] = image_url_for(sl["image_query"], theme)
    return outline


# ---------------------------------------------------------------------------
# STEP 2: Generated placeholder artwork (abstract gradient + shapes)
# so the deck looks designed even without external stock-photo APIs.
# Swap generate_art() for a real stock/image API call if you have a key.
# ---------------------------------------------------------------------------
def generate_art(theme, seed_text, w=1200, h=900):
    c1 = tuple(int(theme["gradient"][0][i:i+2], 16) for i in (0, 2, 4))
    c2 = tuple(int(theme["gradient"][1][i:i+2], 16) for i in (0, 2, 4))
    img = Image.new("RGB", (w, h), c1)
    draw = ImageDraw.Draw(img)
    for y in range(h):
        t = y / h
        r = int(c1[0] + (c2[0]-c1[0])*t)
        g = int(c1[1] + (c2[1]-c1[1])*t)
        b = int(c1[2] + (c2[2]-c1[2])*t)
        draw.line([(0, y), (w, y)], fill=(r, g, b))

    rnd = random.Random(seed_text)
    accent = tuple(int(theme["primary"][i:i+2], 16) for i in (0, 2, 4))
    for _ in range(6):
        x, y = rnd.randint(0, w), rnd.randint(0, h)
        r = rnd.randint(60, 220)
        overlay = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        od = ImageDraw.Draw(overlay)
        od.ellipse([x-r, y-r, x+r, y+r], fill=accent + (35,))
        img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
    img = img.filter(ImageFilter.GaussianBlur(2))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# STEP 3: Flowchart drawn as native PPTX shapes (editable in PowerPoint)
# ---------------------------------------------------------------------------
def add_flowchart(slide, steps, theme, left, top, width, height):
    n = len(steps)
    if n == 0:
        return
    gap = Inches(0.35)
    box_w = (width - gap * (n - 1)) / n
    box_h = min(height, Inches(1.1))
    y = top + (height - box_h) / 2

    boxes = []
    for i, step in enumerate(steps):
        x = left + i * (box_w + gap)
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, box_w, box_h)
        shape.fill.solid()
        shape.fill.fore_color.rgb = hexc(theme["primary"]) if i % 2 == 0 else hexc(theme["secondary"])
        shape.line.fill.background()
        shape.shadow.inherit = False
        tf = shape.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = step
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.color.rgb = hexc(theme["bg"]) if i % 2 == 0 else hexc(theme["text"])
        boxes.append(shape)

    for i in range(n - 1):
        connector = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            boxes[i].left + boxes[i].width, boxes[i].top + boxes[i].height // 2,
            boxes[i+1].left, boxes[i+1].top + boxes[i+1].height // 2,
        )
        connector.line.color.rgb = hexc(theme["muted"])
        connector.line.width = Pt(2)


# ---------------------------------------------------------------------------
# STEP 4: Build the full deck
# ---------------------------------------------------------------------------
def build_pptx(outline, theme_key):
    theme = THEMES.get(theme_key, THEMES["midnight"])
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    def new_slide():
        s = prs.slides.add_slide(blank)
        bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = hexc(theme["bg"])
        bg.line.fill.background()
        bg.shadow.inherit = False
        s.shapes._spTree.remove(bg._element)
        s.shapes._spTree.insert(2, bg._element)
        return s

    def add_text(slide, text, left, top, width, height, size, bold, color, align=PP_ALIGN.LEFT):
        box = slide.shapes.add_textbox(left, top, width, height)
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.name = theme["font"]
        run.font.color.rgb = hexc(color)
        return box

    # --- Title slide ---
    s = new_slide()
    accent = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.25), prs.slide_height)
    accent.fill.solid(); accent.fill.fore_color.rgb = hexc(theme["primary"]); accent.line.fill.background()
    add_text(s, outline.get("title", "Untitled"), Inches(1), Inches(2.8), Inches(11.3), Inches(1.5),
              40, True, theme["text"])
    add_text(s, outline.get("subtitle", ""), Inches(1), Inches(4.0), Inches(11.3), Inches(0.8),
              18, False, theme["muted"])

    # --- Content slides ---
    for sl in outline.get("slides", []):
        stype = sl.get("type", "content")
        title = sl.get("title", "")
        bullets = sl.get("bullets", [])
        needs_img = sl.get("needs_image") and stype != "flowchart"
        needs_flow = sl.get("needs_flowchart") or stype == "flowchart"

        s = new_slide()
        add_text(s, title, Inches(0.7), Inches(0.5), Inches(11.9), Inches(0.9), 28, True, theme["primary"])
        rule = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.7), Inches(1.3), Inches(2.2), Pt(3))
        rule.fill.solid(); rule.fill.fore_color.rgb = hexc(theme["secondary"]); rule.line.fill.background()

        if needs_flow and sl.get("flowchart_steps"):
            add_flowchart(s, sl["flowchart_steps"], theme, Inches(0.7), Inches(2.2), Inches(11.9), Inches(2.0))
            body_top = Inches(4.5)
        else:
            body_top = Inches(1.7)

        if needs_img and not needs_flow:
            img_url = sl.get("image_url")
            try:
                if not img_url:
                    raise ValueError("no image_url on slide")
                img_buf = fetch_image_bytes(img_url)
            except Exception:
                # offline / request failed -> themed placeholder art instead
                img_buf = generate_art(theme, title)
            s.shapes.add_picture(img_buf, Inches(8.3), body_top, width=Inches(4.3), height=Inches(4.6))
            text_width = Inches(7.2)
        else:
            text_width = Inches(11.9)

        if bullets:
            box = s.shapes.add_textbox(Inches(0.7), body_top, text_width, Inches(4.6))
            tf = box.text_frame
            tf.word_wrap = True
            for i, b in enumerate(bullets):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.space_after = Pt(14)
                bullet_run = p.add_run()
                bullet_run.text = "●  "
                bullet_run.font.size = Pt(16)
                bullet_run.font.color.rgb = hexc(theme["secondary"])
                run = p.add_run()
                run.text = b
                run.font.size = Pt(16)
                run.font.color.rgb = hexc(theme["text"])
                run.font.name = theme["font"]

    # --- Closing slide ---
    s = new_slide()
    add_text(s, "Thank You", Inches(1), Inches(3.2), Inches(11.3), Inches(1.2), 36, True, theme["text"], PP_ALIGN.CENTER)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# ROUTES
# ---------------------------------------------------------------------------
@app.route("/api/themes", methods=["GET"])
def get_themes():
    return jsonify({k: {"name": v["name"], "primary": v["primary"], "secondary": v["secondary"], "bg": v["bg"]}
                     for k, v in THEMES.items()})


@app.route("/api/generate-outline", methods=["POST"])
def api_generate_outline():
    data = request.json
    topic = data.get("topic", "").strip()
    theme_key = data.get("theme", "midnight")
    if not topic:
        return jsonify({"error": "Topic is required"}), 400
    try:
        outline = generate_outline(topic, data.get("num_slides", 8), data.get("audience", ""))
        outline = attach_image_urls(outline, theme_key)
        return jsonify(outline)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/generate-pptx", methods=["POST"])
def api_generate_pptx():
    data = request.json
    outline = data.get("outline")
    theme_key = data.get("theme", "midnight")
    if not outline:
        return jsonify({"error": "Outline is required"}), 400
    try:
        buf = build_pptx(outline, theme_key)
        filename = "".join(c for c in outline.get("title", "presentation") if c.isalnum() or c in " -_").strip() or "presentation"
        return send_file(buf, as_attachment=True, download_name=f"{filename}.pptx",
                          mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True, port=5000)
