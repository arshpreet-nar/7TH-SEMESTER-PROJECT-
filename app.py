import base64
import json
import os
from functools import lru_cache
from io import BytesIO
from pathlib import Path

import numpy as np
import tensorflow as tf
from flask import Flask, jsonify, render_template, request, send_from_directory
from PIL import Image


BASE_DIR = Path(__file__).resolve().parent
MODELS_DIR = BASE_DIR / "models"
METADATA_PATH = MODELS_DIR / "metadata.json"
ALLOWED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}
MIN_CONFIDENCE = float(os.environ.get("ECOSORT_MIN_CONFIDENCE", "0.40"))
if not 0 <= MIN_CONFIDENCE <= 1:
    raise ValueError("ECOSORT_MIN_CONFIDENCE must be between 0 and 1.")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 8 * 1024 * 1024


def load_metadata() -> dict:
    if not METADATA_PATH.exists():
        raise FileNotFoundError(f"Metadata not found: {METADATA_PATH}")
    return json.loads(METADATA_PATH.read_text(encoding="utf-8"))


def available_models() -> list[Path]:
    return sorted(MODELS_DIR.glob("*.keras"))


@lru_cache(maxsize=4)
def load_model(model_name: str) -> tf.keras.Model:
    model_path = MODELS_DIR / model_name
    if model_path.parent != MODELS_DIR or not model_path.exists():
        raise FileNotFoundError(f"Model not found: {model_name}")
    return tf.keras.models.load_model(model_path)


def extension_allowed(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def predict_image(image: Image.Image, model_name: str, top_k: int = 5) -> list[dict]:
    metadata = load_metadata()
    class_names = metadata["class_names"]
    image_size = tuple(metadata["image_size"])

    resized = image.convert("RGB").resize(image_size)
    array = tf.keras.utils.img_to_array(resized)
    probabilities = load_model(model_name).predict(np.expand_dims(array, axis=0), verbose=0)[0]
    top_indices = probabilities.argsort()[-top_k:][::-1]

    return [
        {
            "label": class_names[index],
            "probability": float(probabilities[index]),
            "percent": round(float(probabilities[index]) * 100, 2),
        }
        for index in top_indices
    ]


def prediction_is_confident(predictions: list[dict]) -> bool:
    return bool(predictions and predictions[0]["probability"] >= MIN_CONFIDENCE)


def image_to_data_url(image: Image.Image) -> str:
    buffer = BytesIO()
    image.convert("RGB").save(buffer, format="JPEG", quality=88)
    encoded = base64.b64encode(buffer.getvalue()).decode("ascii")
    return f"data:image/jpeg;base64,{encoded}"


def image_from_data_url(data_url: str) -> Image.Image:
    if not data_url.startswith("data:image/") or "," not in data_url:
        raise ValueError("Camera frame must be a valid image data URL.")

    encoded = data_url.split(",", 1)[1]
    if len(encoded) > 11_000_000:
        raise ValueError("Camera frame is too large.")

    try:
        return Image.open(BytesIO(base64.b64decode(encoded, validate=True)))
    except Exception as exc:
        raise ValueError("Camera frame could not be read.") from exc


@app.route("/", methods=["GET", "POST"])
def index():
    metadata = load_metadata()
    model_paths = available_models()
    model_names = [path.name for path in model_paths]
    selected_model = request.form.get("model") or (
        "best_garbage_classifier.keras"
        if "best_garbage_classifier.keras" in model_names
        else model_names[0]
        if model_names
        else ""
    )

    result = None
    error = None

    if request.method == "POST":
        uploaded_file = request.files.get("image")
        if not uploaded_file or not uploaded_file.filename:
            error = "Choose an image before running prediction."
        elif not extension_allowed(uploaded_file.filename):
            error = "Use a JPG, PNG, BMP, or WEBP image."
        elif selected_model not in model_names:
            error = "Choose one of the available trained models."
        else:
            try:
                image = Image.open(uploaded_file.stream)
                predictions = predict_image(image, selected_model)
                result = {
                    "filename": uploaded_file.filename,
                    "image": image_to_data_url(image),
                    "predictions": predictions,
                    "top_prediction": predictions[0],
                    "confident": prediction_is_confident(predictions),
                    "threshold_percent": round(MIN_CONFIDENCE * 100),
                }
            except Exception as exc:
                error = f"Could not classify this image: {exc}"

    return render_template(
        "index.html",
        metadata=metadata,
        models=model_names,
        selected_model=selected_model,
        result=result,
        error=error,
    )


@app.post("/api/predict-frame")
def predict_frame():
    payload = request.get_json(silent=True) or {}
    model_name = payload.get("model", "")
    model_names = [path.name for path in available_models()]

    if model_name not in model_names:
        return jsonify(error="Choose one of the available trained models."), 400

    frame = payload.get("image")
    if not isinstance(frame, str):
        return jsonify(error="Send an image captured from the camera."), 400

    try:
        image = image_from_data_url(frame)
        predictions = predict_image(image, model_name)
    except Exception as exc:
        return jsonify(error=f"Could not classify this camera frame: {exc}"), 400

    return jsonify(
        top_prediction=predictions[0],
        predictions=predictions,
        confident=prediction_is_confident(predictions),
        threshold_percent=round(MIN_CONFIDENCE * 100),
    )


@app.get("/service-worker.js")
def service_worker():
    return send_from_directory(
        app.static_folder,
        "service-worker.js",
        mimetype="application/javascript",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="127.0.0.1", port=port, debug=True, use_reloader=False)
