import argparse
import json
from pathlib import Path

import numpy as np
import tensorflow as tf


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Predict the garbage class for one image.")
    parser.add_argument("image", type=Path, help="Path to the image to classify.")
    parser.add_argument("--model", type=Path, default=Path("models/best_garbage_classifier.keras"))
    parser.add_argument("--metadata", type=Path, default=Path("models/metadata.json"))
    parser.add_argument("--top-k", type=int, default=3)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not args.image.exists():
        raise FileNotFoundError(f"Image not found: {args.image}")
    if not args.model.exists():
        raise FileNotFoundError(f"Model not found: {args.model}")
    if not args.metadata.exists():
        raise FileNotFoundError(f"Metadata not found: {args.metadata}")

    metadata = json.loads(args.metadata.read_text(encoding="utf-8"))
    class_names = metadata["class_names"]
    image_size = tuple(metadata["image_size"])

    model = tf.keras.models.load_model(args.model)
    image = tf.keras.utils.load_img(args.image, target_size=image_size)
    array = tf.keras.utils.img_to_array(image)
    batch = np.expand_dims(array, axis=0)

    probabilities = model.predict(batch, verbose=0)[0]
    top_indices = probabilities.argsort()[-args.top_k :][::-1]

    for index in top_indices:
        print(f"{class_names[index]}: {probabilities[index]:.4f}")


if __name__ == "__main__":
    main()
