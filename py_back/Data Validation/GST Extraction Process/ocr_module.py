import cv2
import numpy as np
import torch
from transformers import TrOCRProcessor, VisionEncoderDecoderModel
import easyocr
from PIL import Image

processor = None
model = None
easyocr_reader = None
_trocr_init_error = None
_easyocr_init_error = None


def _ensure_trocr_components():
    global processor, model, _trocr_init_error
    if processor is not None and model is not None:
        return processor, model
    if _trocr_init_error is not None:
        raise RuntimeError("TrOCR initialization previously failed") from _trocr_init_error

    try:
        processor = TrOCRProcessor.from_pretrained("microsoft/trocr-large-printed")
        model = VisionEncoderDecoderModel.from_pretrained(
            "microsoft/trocr-large-printed",
            low_cpu_mem_usage=False,
        )
        model.to("cpu")
        model.eval()
        return processor, model
    except Exception as exc:
        _trocr_init_error = exc
        raise


def _ensure_easyocr_reader():
    global easyocr_reader, _easyocr_init_error
    if easyocr_reader is not None:
        return easyocr_reader
    if _easyocr_init_error is not None:
        raise RuntimeError("EasyOCR initialization previously failed") from _easyocr_init_error

    try:
        easyocr_reader = easyocr.Reader(['en'], gpu=False)  # Set gpu=True if you have CUDA
        return easyocr_reader
    except Exception as exc:
        _easyocr_init_error = exc
        raise


def _trocr_decode(pil_img):
    """Run TrOCR in CPU-only inference mode and return decoded strings."""
    processor_ref, model_ref = _ensure_trocr_components()
    with torch.inference_mode():
        pixel_values = processor_ref(images=pil_img, return_tensors="pt").pixel_values.to("cpu")
        generated_ids = model_ref.generate(pixel_values, max_new_tokens=12)
    return processor_ref.batch_decode(generated_ids, skip_special_tokens=True)


def tr_ocr_image_from_array(image_array):
    """
    Best for single line captcha text
    Input: numpy array
    Output: List of detected text strings
    """
    img = Image.fromarray(image_array).convert("RGB")
    return _trocr_decode(img)


def tr_ocr_split_horizontally(image_array):
    """
    Best for stacked/multi-line captcha text
    Input: numpy array
    Output: List of detected text strings (one per line)
    """
    # Split image horizontally into two parts
    height, width = image_array.shape[:2]
    mid_height = height // 2
    
    # Top half
    top_half = image_array[:mid_height, :]
    top_pil = Image.fromarray(top_half).convert("RGB")
    
    # Bottom half
    bottom_half = image_array[mid_height:, :]
    bottom_pil = Image.fromarray(bottom_half).convert("RGB")
    
    all_texts = []
    
    # Process each half
    for half_img in [top_pil, bottom_pil]:
        text = _trocr_decode(half_img)[0]
        if text.strip():  # Only add non-empty text
            all_texts.append(text.strip())
    
    return all_texts


def easy_ocr_image_from_array(image_array):
    """
    Fallback method using EasyOCR + TrOCR combination
    Best for complex/unclear captchas
    Input: numpy array
    Output: List of detected text strings
    """
    # Detect text regions
    reader = _ensure_easyocr_reader()
    results = reader.readtext(image_array)
    
    all_texts = []
    
    for (bbox, _, confidence) in results:
        if confidence > 0.5:  # Filter by confidence
            # Extract bounding box coordinates
            x_min = int(min([p[0] for p in bbox]))
            y_min = int(min([p[1] for p in bbox]))
            x_max = int(max([p[0] for p in bbox]))
            y_max = int(max([p[1] for p in bbox]))
            
            # Crop the text region
            cropped = image_array[y_min:y_max, x_min:x_max]
            cropped_pil = Image.fromarray(cropped).convert("RGB")
            
            # Apply TrOCR
            try:
                text = _trocr_decode(cropped_pil)[0]
            except Exception:
                # Final fallback: use EasyOCR's own detected text.
                text = ""
            if text.strip():
                all_texts.append(text.strip())
    
    return all_texts


def special_character_remover(raw_output):
    """
    Remove any special characters from detected text
    Input: str
    Output: str (cleaned)
    """
    if not raw_output:
        return ""
    
    import re
    cleaned = re.sub(r'[^A-Za-z0-9]', '', raw_output)
    return cleaned


def extract_text(cropped_image):
    # Method 1: Try TrOCR for single line
    try:
        extracted_text = tr_ocr_image_from_array(cropped_image)
    except Exception:
        extracted_text = []
    temp_text = " ".join(extracted_text)
    
    if len(temp_text) >= 3:  # If result is good enough
        return extracted_text
    
    # Method 2: Try splitting horizontally (for stacked text)
    try:
        extracted_text = tr_ocr_split_horizontally(cropped_image)
    except Exception:
        extracted_text = []
    temp_text = " ".join(extracted_text)
    
    if len(temp_text) >= 3:  # If result is good enough
        return extracted_text
    
    # Method 3: Fallback to EasyOCR
    try:
        extracted_text = easy_ocr_image_from_array(cropped_image)
    except Exception:
        extracted_text = []
    return extracted_text
