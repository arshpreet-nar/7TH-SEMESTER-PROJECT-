import argparse
import json
from pathlib import Path

import matplotlib.pyplot as plt
import tensorflow as tf


AUTOTUNE = tf.data.AUTOTUNE


def count_images(data_dir: Path, class_names: list[str]) -> dict[str, int]:
    counts = {}
    for class_name in class_names:
        class_dir = data_dir / class_name
        counts[class_name] = sum(
            1
            for path in class_dir.iterdir()
            if path.is_file() and path.suffix.lower() in {".jpg", ".jpeg", ".png", ".bmp", ".webp"}
        )
    return counts


def make_class_weights(counts: dict[str, int], class_names: list[str]) -> dict[int, float]:
    total = sum(counts.values())
    num_classes = len(class_names)
    return {
        index: total / (num_classes * max(counts[class_name], 1))
        for index, class_name in enumerate(class_names)
    }


def prepare_dataset(dataset: tf.data.Dataset) -> tf.data.Dataset:
    return dataset.prefetch(AUTOTUNE)


def build_model(num_classes: int, image_size: tuple[int, int], dropout: float) -> tf.keras.Model:
    inputs = tf.keras.Input(shape=(*image_size, 3))
    augmentation = tf.keras.Sequential(
        [
            tf.keras.layers.RandomFlip("horizontal"),
            tf.keras.layers.RandomRotation(0.08),
            tf.keras.layers.RandomZoom(0.12),
            tf.keras.layers.RandomContrast(0.1),
        ],
        name="augmentation",
    )
    x = augmentation(inputs)
    x = tf.keras.applications.mobilenet_v2.preprocess_input(x)
    base_model = tf.keras.applications.MobileNetV2(
        input_shape=(*image_size, 3),
        include_top=False,
        weights="imagenet",
    )
    base_model.trainable = False

    x = base_model(x, training=False)
    x = tf.keras.layers.GlobalAveragePooling2D()(x)
    x = tf.keras.layers.Dropout(dropout)(x)
    outputs = tf.keras.layers.Dense(num_classes, activation="softmax")(x)

    model = tf.keras.Model(inputs, outputs)
    model.compile(
        optimizer=tf.keras.optimizers.Adam(),
        loss="sparse_categorical_crossentropy",
        metrics=["accuracy"],
    )
    return model


def plot_history(history: tf.keras.callbacks.History, output_path: Path) -> None:
    metrics = history.history
    epochs = range(1, len(metrics["loss"]) + 1)

    plt.figure(figsize=(10, 4))
    plt.subplot(1, 2, 1)
    plt.plot(epochs, metrics["accuracy"], label="train")
    plt.plot(epochs, metrics["val_accuracy"], label="validation")
    plt.title("Accuracy")
    plt.xlabel("Epoch")
    plt.legend()

    plt.subplot(1, 2, 2)
    plt.plot(epochs, metrics["loss"], label="train")
    plt.plot(epochs, metrics["val_loss"], label="validation")
    plt.title("Loss")
    plt.xlabel("Epoch")
    plt.legend()

    plt.tight_layout()
    plt.savefig(output_path, dpi=150)
    plt.close()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Train a garbage image classifier.")
    parser.add_argument("--data-dir", type=Path, default=Path("garbage_classification"))
    parser.add_argument("--output-dir", type=Path, default=Path("models"))
    parser.add_argument("--image-size", type=int, default=224)
    parser.add_argument("--batch-size", type=int, default=32)
    parser.add_argument("--epochs", type=int, default=12)
    parser.add_argument("--fine-tune-epochs", type=int, default=5)
    parser.add_argument("--validation-split", type=float, default=0.2)
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--dropout", type=float, default=0.25)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not args.data_dir.exists():
        raise FileNotFoundError(f"Dataset folder not found: {args.data_dir}")

    args.output_dir.mkdir(parents=True, exist_ok=True)
    image_size = (args.image_size, args.image_size)

    train_ds = tf.keras.utils.image_dataset_from_directory(
        args.data_dir,
        validation_split=args.validation_split,
        subset="training",
        seed=args.seed,
        image_size=image_size,
        batch_size=args.batch_size,
        label_mode="int",
    )
    val_ds = tf.keras.utils.image_dataset_from_directory(
        args.data_dir,
        validation_split=args.validation_split,
        subset="validation",
        seed=args.seed,
        image_size=image_size,
        batch_size=args.batch_size,
        label_mode="int",
    )

    class_names = train_ds.class_names
    counts = count_images(args.data_dir, class_names)
    class_weights = make_class_weights(counts, class_names)

    metadata = {
        "class_names": class_names,
        "image_size": list(image_size),
        "class_counts": counts,
        "class_weights": class_weights,
    }
    (args.output_dir / "metadata.json").write_text(json.dumps(metadata, indent=2), encoding="utf-8")

    train_ds = prepare_dataset(train_ds)
    val_ds = prepare_dataset(val_ds)

    model = build_model(len(class_names), image_size, args.dropout)
    checkpoint_path = args.output_dir / "best_garbage_classifier.keras"
    callbacks = [
        tf.keras.callbacks.ModelCheckpoint(
            checkpoint_path,
            monitor="val_accuracy",
            save_best_only=True,
            mode="max",
        ),
        tf.keras.callbacks.EarlyStopping(
            monitor="val_loss",
            patience=4,
            restore_best_weights=True,
        ),
        tf.keras.callbacks.ReduceLROnPlateau(
            monitor="val_loss",
            factor=0.3,
            patience=2,
            min_lr=1e-6,
        ),
    ]

    print("Classes:", ", ".join(class_names))
    print("Image counts:", counts)
    print("Class weights:", class_weights)

    history = model.fit(
        train_ds,
        validation_data=val_ds,
        epochs=args.epochs,
        class_weight=class_weights,
        callbacks=callbacks,
    )

    if args.fine_tune_epochs > 0:
        base_model = next(layer for layer in model.layers if isinstance(layer, tf.keras.Model))
        base_model.trainable = True
        for layer in base_model.layers[:-30]:
            layer.trainable = False

        model.compile(
            optimizer=tf.keras.optimizers.Adam(learning_rate=1e-5),
            loss="sparse_categorical_crossentropy",
            metrics=["accuracy"],
        )
        history_fine = model.fit(
            train_ds,
            validation_data=val_ds,
            epochs=args.epochs + args.fine_tune_epochs,
            initial_epoch=len(history.history["loss"]),
            class_weight=class_weights,
            callbacks=callbacks,
        )
        for key, values in history_fine.history.items():
            history.history.setdefault(key, []).extend(values)

    final_model_path = args.output_dir / "garbage_classifier_final.keras"
    model.save(final_model_path)
    plot_history(history, args.output_dir / "training_history.png")

    loss, accuracy = model.evaluate(val_ds, verbose=0)
    print(f"Validation accuracy: {accuracy:.4f}")
    print(f"Validation loss: {loss:.4f}")
    print(f"Saved best model to: {checkpoint_path}")
    print(f"Saved final model to: {final_model_path}")
    print(f"Saved metadata to: {args.output_dir / 'metadata.json'}")


if __name__ == "__main__":
    main()
