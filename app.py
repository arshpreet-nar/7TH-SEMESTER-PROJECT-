import base64
import json
import os
from functools import lru_cache
from io import BytesIO
from pathlib import Path

import numpy as np
import tensorflow as tf
from flask import Flask, render_template, request
from PIL import Image


BASE_DIR = Path(__file__).resolve().parent
MODELS_DIR = BASE_DIR / "models"
METADATA_PATH = MODELS_DIR / "metadata.json"
ALLOWED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}

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


def image_to_data_url(image: Image.Image) -> str:
    buffer = BytesIO()
    image.convert("RGB").save(buffer, format="JPEG", quality=88)
    encoded = base64.b64encode(buffer.getvalue()).decode("ascii")
    return f"data:image/jpeg;base64,{encoded}"


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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="127.0.0.1", port=port, debug=True, use_reloader=False)
