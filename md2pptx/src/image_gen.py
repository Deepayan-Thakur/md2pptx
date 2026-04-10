import os
import hashlib
import io
from typing import Optional

try:
    from huggingface_hub import InferenceClient
except ImportError:
    InferenceClient = None

def _get_cache_dir():
    cache_dir = os.path.join(os.path.dirname(__file__), "..", "..", "md2pptx", "outputs", ".cache")
    os.makedirs(cache_dir, exist_ok=True)
    return cache_dir

def _hash_prompt(prompt: str) -> str:
    return hashlib.md5(prompt.encode('utf-8')).hexdigest()

def generate_slide_asset(title: str, doc_title: str) -> Optional[bytes]:
    """
    Generates a corporate presentation slide asset using Hugging Face Text-to-Image API.
    Uses local disk caching to prevent API rate limits.
    """
    api_key = os.getenv("HUGGINGFACE_API_KEY")
    if not api_key or InferenceClient is None:
        return None

    # Enhanced dynamic prompt specifically designed for slides
    prompt = f"A professional corporate presentation slide illustration for '{title}', related to '{doc_title}'. Minimalist vector art style, deep corporate blues and reds, clean white background, high quality, professional business isometric."
    
    # Check cache first
    prompt_hash = _hash_prompt(prompt)
    cache_path = os.path.join(_get_cache_dir(), f"{prompt_hash}.jpg")
    
    if os.path.exists(cache_path):
        print(f"   [Cache Hit] Image asset for: {title[:30]}...")
        try:
            with open(cache_path, "rb") as f:
                return f.read()
        except:
            pass
            
    print(f"   [ImageGen] Generating asset via HF for: {title[:30]}...")
    
    # Using FLUX.1-schnell as it is incredible at following aesthetic commands quickly
    try:
        client = InferenceClient("black-forest-labs/FLUX.1-schnell", token=api_key)
        image = client.text_to_image(prompt, width=1024, height=576) # 16:9
        
        # Save to buffer
        stream = io.BytesIO()
        image.save(stream, format="JPEG", quality=85)
        img_bytes = stream.getvalue()
        
        # Save to cache
        with open(cache_path, "wb") as f:
            f.write(img_bytes)
            
        return img_bytes
        
    except Exception as e:
        print(f"   ⚠ [ImageGen Failed] {e}")
        return None
