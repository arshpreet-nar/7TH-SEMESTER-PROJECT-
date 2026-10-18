# EcoSort AI — Garbage Classification

EcoSort AI is an image classification project that identifies common waste categories. It includes a trained TensorFlow model, a Flask web app with image upload and camera capture, the source dataset, and project presentation materials.

## Project contents

| Path | Purpose |
| --- | --- |
| `app.py` | Flask application and model inference |
| `templates/` and `static/` | Web interface, styles, app icon, and offline support |
| `models/` | Trained Keras models, class metadata, and training history |
| `garbage_classification/` | Labeled image dataset and sample test images |
| `garbage_classification_model.ipynb` | Model development notebook |
| `Report/` | Final project report and college presentation |
| `Output/` | Latest presentation output |
| `project_archive/presentation-build/` | Preserved presentation source scripts and intermediate build assets |

## Waste categories

The training dataset has 12 class folders: battery, biological, brown-glass, cardboard, clothes, green-glass, metal, paper, plastic, shoes, trash, and white-glass. The `testing/` folder contains sample images and `expected_labels.csv` for reference.

## Requirements

- Python 3.11 (the deployment configuration pins 3.11.9)
- Packages listed in `requirements.txt`
- The included model files in `models/`

## Run locally

In PowerShell, from the project root:

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python app.py
```

Open <http://127.0.0.1:5000>. The interface supports uploading an image or capturing a still image with a camera. Camera access requires `localhost` during local development or HTTPS after deployment. Set `ECOSORT_MIN_CONFIDENCE` to a value between `0` and `1` to adjust the low-confidence warning threshold (default: `0.40`).

## Deployment

The repository includes `render.yaml`, `runtime.txt`, `.python-version`, and `Procfile` for Render deployment. The service uses:

- Build: `pip install -r requirements.txt`
- Start: `gunicorn --bind 0.0.0.0:$PORT app:app`
- Python: `3.11.9`

## Deliverables

- [Project report](Report/EcoSort_AI_Professional_Project_Report_50_Page_Edition.pdf)
- [Final presentation](Output/EcoSort_AI_Final_College_Presentation_with_logo_fixed.pptx)

The project archive retains build scripts and intermediate visuals that may be useful for revising those deliverables. Python caches and temporary lock files are excluded from the project.
