import os
import hashlib
import io
import math
import random
import requests
from typing import Optional

try:
    from PIL import Image, ImageDraw, ImageFilter
    _PIL_AVAILABLE = True
except ImportError:
    _PIL_AVAILABLE = False


def _get_cache_dir() -> str:
    cache_dir = os.path.join(os.path.dirname(__file__), "..", "..", "md2pptx", "outputs", ".cache")
    os.makedirs(cache_dir, exist_ok=True)
    return cache_dir


def _hash_prompt(prompt: str) -> str:
    return hashlib.md5(prompt.encode("utf-8")).hexdigest()


def _get_unsplash_image(query: str) -> Optional[bytes]:
    """Fetch high-quality professional stock photo from Unsplash (Free Tier: 50 req/hr)."""
    access_key = os.getenv("UNSPLASH_ACCESS_KEY")
    if not access_key:
        return None
        
    try:
        url = "https://api.unsplash.com/photos/random"
        headers = {"Authorization": f"Client-ID {access_key}"}
        params = {
            "query": query,
            "orientation": "landscape",
            "content_filter": "high"
        }
        resp = requests.get(url, headers=headers, params=params, timeout=10)
        if resp.status_code == 200:
            img_url = resp.json()["urls"]["regular"]
            img_resp = requests.get(img_url, timeout=15)
            if img_resp.status_code == 200:
                return img_resp.content
    except Exception as e:
        print(f"   [Unsplash Error] {e}")
    return None


def _get_pollinations_image(prompt: str) -> Optional[bytes]:
    """Fetch free AI-generated art from Pollinations.ai (No key required)."""
    try:
        # Encode prompt for URL
        from urllib.parse import quote
        encoded = quote(prompt)
        url = f"https://image.pollinations.ai/prompt/{encoded}?width=1024&height=576&nologo=true&seed={random.randint(1, 99999)}"
        resp = requests.get(url, timeout=25)
        if resp.status_code == 200:
            return resp.content
    except Exception as e:
        print(f"   [Pollinations Error] {e}")
    return None


def _build_prompt_and_keywords(title: str, doc_title: str, is_mascot: bool = False) -> (str, str):
    """Build both an AI prompt and search keywords for the slide context."""
    combined = (doc_title + " " + title).lower()
    
    # Base prompts
    if is_mascot:
        keywords = "3D character robot standing mascot white background"
        prompt = "high-quality 3D render of a friendly sleek white and blue robot mascot character, standing in a professional pose, clean white background, futuristic design, 8k, Octane render"
    elif any(k in combined for k in ["ai", "digital", "tech", "data", "bubble", "mechanism"]):
        keywords = "futuristic technology red neural network abstract"
        prompt = "ultra-modern futuristic AI technology concept art, glowing red and white neural network connections, abstract digital background, 8k, photorealistic"
    elif any(k in combined for k in ["security", "cyber", "threat", "breach"]):
        keywords = "cybersecurity digital shield privacy"
        prompt = "high-tech cybersecurity illustration, glowing digital shields, dark background, sleek corporate render"
    elif any(k in combined for k in ["finance", "revenue", "growth", "market", "money"]):
        keywords = "stock market graph growth investment"
        prompt = "business growth and investment concept art, rising bar charts with glowing lines, deep navy background"
    elif any(k in combined for k in ["strategy", "plan", "roadmap", "objective"]):
        keywords = "business strategy roadmap success"
        prompt = "corporate strategy concept, abstract roadmap to success, professional minimal aesthetic"
    else:
        keywords = "corporate business professional office"
        prompt = "premium clean corporate business illustration, abstract minimal geometric shapes, sleek professional aesthetic"
        
    return prompt, keywords


def _generate_fallback_image(title: str, width: int = 1024, height: int = 576) -> Optional[bytes]:
    """Programmatic corporate background image using PIL if APIs fail."""
    if not _PIL_AVAILABLE:
        return None
    bg_color = (20, 25, 50)
    img = Image.new("RGB", (width, height), bg_color)
    draw = ImageDraw.Draw(img)
    # Simple grid pattern
    for x in range(0, width, 100): draw.line([(x, 0), (x, height)], fill=(40, 50, 80))
    for y in range(0, height, 100): draw.line([(0, y), (width, y)], fill=(40, 50, 80))
    # Horizontal accent bar
    draw.rectangle([(0, height - 30), (width, height)], fill=(239, 68, 68)) # Accenture Red
    stream = io.BytesIO()
    img.save(stream, format="JPEG", quality=85)
    return stream.getvalue()


def generate_slide_asset(title: str, doc_title: str, is_mascot: bool = False) -> Optional[bytes]:
    """
    Multi-source free image engine with mascot support:
    1. Unsplash (Professional Photos)
    2. Pollinations (Free AI Art / Mascots)
    3. Programmatic Fallback
    """
    ai_prompt, search_keywords = _build_prompt_and_keywords(title, doc_title, is_mascot)
    
    # Use different cache key for mascots
    cache_id = f"{_hash_prompt(search_keywords)}{'_mascot' if is_mascot else ''}"
    cache_path = os.path.join(_get_cache_dir(), f"{cache_id}.jpg")

    if os.path.exists(cache_path):
        with open(cache_path, "rb") as f:
            data = f.read()
        if data and len(data) > 512:  # sanity check: valid image
            return data
        else:
            os.remove(cache_path)  # purge corrupt cache

    # 1. Try Unsplash (Skip for mascots as AI art is better for custom 3D characters)
    if not is_mascot and os.getenv("UNSPLASH_ACCESS_KEY"):
        print(f"   [Unsplash] Fetching asset for: {title[:30]}...")
        img_bytes = _get_unsplash_image(search_keywords)
        if img_bytes and len(img_bytes) > 512:
            with open(cache_path, "wb") as f: f.write(img_bytes)
            return img_bytes

    # 2. Try Pollinations AI Art
    print(f"   [Pollinations] Generating {'Mascot' if is_mascot else 'AI art'} for: {title[:30]}...")
    img_bytes = _get_pollinations_image(ai_prompt)
    if img_bytes and len(img_bytes) > 512:
        with open(cache_path, "wb") as f: f.write(img_bytes)
        return img_bytes

    # 3. Last Resort Fallback
    print(f"   [Fallback] Generating geometric art for: {title[:30]}...")
    return _generate_fallback_image(title)
