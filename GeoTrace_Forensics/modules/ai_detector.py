# modules/ai_detector.py — Teammate 1: AI/ML Image Manipulation Detector

import hashlib
import math
import os
from pathlib import Path

import numpy as np
from PIL import Image, ImageChops, ImageEnhance, ImageFilter


# ─────────────────────────────────────────────
# 1. ERROR LEVEL ANALYSIS (ELA)
# The core AI detection technique used in real forensics.
# Idea: re-save image at low quality, compare to original.
# Edited regions compress differently → show as bright spots.
# ─────────────────────────────────────────────
def run_ela(image_path, quality=75):
    """
    Run Error Level Analysis on an image.
    Returns (ela_image, max_error, mean_error, suspicious_score)
    """
    image_path = Path(image_path)
    original = Image.open(image_path).convert("RGB")

    # Re-save at lower quality to a temp file
    temp_path = image_path.parent / f"_ela_temp_{image_path.stem}.jpg"
    original.save(str(temp_path), "JPEG", quality=quality)

    # Reload the re-saved version
    recompressed = Image.open(temp_path).convert("RGB")

    # Calculate pixel-by-pixel difference
    ela_image = ImageChops.difference(original, recompressed)

    # Amplify the difference so it's visible
    extrema = ela_image.getextrema()
    max_diff = max([ex[1] for ex in extrema]) or 1
    scale = 255.0 / max_diff
    ela_image = ImageEnhance.Brightness(ela_image).enhance(scale)

    # Get error statistics
    ela_array = np.array(ela_image)
    max_error  = float(np.max(ela_array))
    mean_error = float(np.mean(ela_array))

    # Clean up temp file
    try:
        os.remove(temp_path)
    except Exception:
        pass

    # Score: 0.0 = clean, 1.0 = highly suspicious
    # Real forensic threshold: mean error > 15 is suspicious
    suspicious_score = min(mean_error / 25.0, 1.0)

    return ela_image, max_error, mean_error, suspicious_score


# ─────────────────────────────────────────────
# 2. NOISE ANALYSIS
# Authentic photos have consistent sensor noise.
# Edited/AI images often have uneven or missing noise.
# ─────────────────────────────────────────────
def analyze_noise(image_path):
    """
    Analyze image noise consistency across regions.
    Returns (noise_score, noise_variance, verdict)
    """
    img = Image.open(image_path).convert("L")  # grayscale

    # Apply edge-detection to isolate noise
    edges = img.filter(ImageFilter.FIND_EDGES)
    noise_array = np.array(edges, dtype=float)

    # Divide image into 4 quadrants, compare noise levels
    h, w = noise_array.shape
    quadrants = [
        noise_array[:h//2, :w//2],   # top-left
        noise_array[:h//2, w//2:],   # top-right
        noise_array[h//2:, :w//2],   # bottom-left
        noise_array[h//2:, w//2:],   # bottom-right
    ]

    quad_means = [float(np.mean(q)) for q in quadrants]
    noise_variance = float(np.var(quad_means))

    # High variance = inconsistent noise = possible manipulation
    # Threshold calibrated on real images
    noise_score = min(noise_variance / 500.0, 1.0)

    if noise_score < 0.3:
        verdict = "Consistent noise (authentic)"
    elif noise_score < 0.6:
        verdict = "Slightly inconsistent noise (possible edit)"
    else:
        verdict = "Highly inconsistent noise (likely manipulated)"

    return noise_score, noise_variance, verdict


# ─────────────────────────────────────────────
# 3. METADATA CONSISTENCY CHECK
# Real camera images have matching metadata.
# Edited images often have mismatches.
# ─────────────────────────────────────────────
def check_metadata_consistency(exif_data):
    """
    Cross-check EXIF fields for internal consistency.
    Returns list of inconsistency findings.
    """
    findings = []

    if not exif_data:
        findings.append({
            "check"   : "EXIF Presence",
            "result"  : "FAIL",
            "detail"  : "No EXIF data found — metadata may have been stripped."
        })
        return findings

    # Check 1: Software tag (editing software leaves a trace)
    software = exif_data.get("software", "")
    editing_tools = ["photoshop", "gimp", "lightroom", "affinity",
                     "snapseed", "facetune", "canva", "pixlr"]
    if software and any(tool in software.lower() for tool in editing_tools):
        findings.append({
            "check"  : "Editing Software",
            "result" : "WARN",
            "detail" : f"Editing software detected: {software}"
        })

    # Check 2: Date consistency
    date_taken    = exif_data.get("date_taken", "")
    date_modified = exif_data.get("date_modified", "")
    if date_taken and date_modified and date_taken != date_modified:
        findings.append({
            "check"  : "Date Consistency",
            "result" : "WARN",
            "detail" : f"Date taken ({date_taken}) differs from date modified ({date_modified})"
        })

    # Check 3: GPS vs camera make — phones always have GPS
    camera_make = exif_data.get("camera_make", "")
    has_gps     = exif_data.get("latitude") is not None
    phone_brands = ["apple", "samsung", "xiaomi", "huawei",
                    "google", "oneplus", "oppo", "vivo"]
    if camera_make and any(brand in camera_make.lower() for brand in phone_brands):
        if not has_gps:
            findings.append({
                "check"  : "GPS Consistency",
                "result" : "WARN",
                "detail" : f"Phone camera ({camera_make}) detected but no GPS — location may have been stripped."
            })

    # Check 4: Missing thumbnail (common after heavy editing)
    if not exif_data.get("thumbnail"):
        findings.append({
            "check"  : "Thumbnail",
            "result" : "INFO",
            "detail" : "No embedded thumbnail — image may have been re-exported."
        })

    if not findings:
        findings.append({
            "check"  : "Metadata",
            "result" : "PASS",
            "detail" : "No metadata inconsistencies detected."
        })

    return findings


# ─────────────────────────────────────────────
# 4. FINAL VERDICT
# Combine all scores into one verdict
# ─────────────────────────────────────────────
def get_manipulation_verdict(ela_score, noise_score, metadata_findings):
    """
    Combine all signals into a final manipulation verdict.
    Returns (verdict_label, confidence_pct, color)
    """
    # Count metadata warnings
    warn_count = sum(1 for f in metadata_findings if f["result"] in ("WARN", "FAIL"))

    # Weighted score
    metadata_score = min(warn_count / 3.0, 1.0)
    combined = (ela_score * 0.5) + (noise_score * 0.3) + (metadata_score * 0.2)

    confidence_pct = round(combined * 100, 1)

    if combined < 0.25:
        return "✅ Likely Authentic",    confidence_pct, "green"
    elif combined < 0.5:
        return "⚠️  Possibly Modified",  confidence_pct, "orange"
    elif combined < 0.75:
        return "🚨 Probably Manipulated", confidence_pct, "red"
    else:
        return "🚨 Almost Certainly Manipulated", confidence_pct, "darkred"


# ─────────────────────────────────────────────
# 5. MAIN FUNCTION — call this from the UI
# ─────────────────────────────────────────────
def analyze_image_authenticity(image_path, exif_data=None):
    """
    Full AI authenticity analysis.
    Returns a dict with all results — plug this into the UI tab.

    Usage:
        result = analyze_image_authenticity("photo.jpg", exif_data=row)
    """
    result = {
        "image_path"        : str(image_path),
        "ela_score"         : 0.0,
        "ela_mean_error"    : 0.0,
        "ela_max_error"     : 0.0,
        "ela_image"         : None,
        "noise_score"       : 0.0,
        "noise_variance"    : 0.0,
        "noise_verdict"     : "",
        "metadata_findings" : [],
        "verdict"           : "",
        "confidence_pct"    : 0.0,
        "verdict_color"     : "gray",
        "error"             : None,
    }

    try:
        # Run ELA
        ela_img, max_err, mean_err, ela_score = run_ela(image_path)
        result["ela_image"]      = ela_img
        result["ela_score"]      = round(ela_score, 3)
        result["ela_mean_error"] = round(mean_err, 2)
        result["ela_max_error"]  = round(max_err, 2)

        # Run noise analysis
        noise_score, noise_var, noise_verdict = analyze_noise(image_path)
        result["noise_score"]   = round(noise_score, 3)
        result["noise_variance"]= round(noise_var, 2)
        result["noise_verdict"] = noise_verdict

        # Metadata consistency
        findings = check_metadata_consistency(exif_data or {})
        result["metadata_findings"] = findings

        # Final verdict
        verdict, confidence, color = get_manipulation_verdict(
            ela_score, noise_score, findings
        )
        result["verdict"]       = verdict
        result["confidence_pct"]= confidence
        result["verdict_color"] = color

    except Exception as e:
        result["error"]  = str(e)
        result["verdict"]= "❌ Analysis failed"

    return result