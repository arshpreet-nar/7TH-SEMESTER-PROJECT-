# Garbage Classification Deep Learning Project

This project trains an image classifier on the `garbage_classification` folder.
The dataset is arranged as one folder per class, so TensorFlow can load it
directly.

## Dataset Classes

Current class counts:

| Class | Images |
| --- | ---: |
| battery | 945 |
| biological | 985 |
| brown-glass | 607 |
| cardboard | 891 |
| clothes | 5325 |
| green-glass | 629 |
| metal | 769 |
| paper | 1050 |
| plastic | 865 |
| shoes | 1977 |
| trash | 697 |
| white-glass | 775 |

## Setup

Install Python 3.10 or 3.11 first. Then run:

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

If your Python version is 3.10, replace `py -3.11` with `py -3.10`.

## Train

```powershell
python train_model.py
```

The script uses MobileNetV2 transfer learning, validation split, early stopping,
learning-rate reduction, and class weights because the dataset is imbalanced.
Outputs are saved in `models/`:

- `best_garbage_classifier.keras`
- `garbage_classifier_final.keras`
- `metadata.json`
- `training_history.png`

For a quicker test run:

```powershell
python train_model.py --epochs 2 --fine-tune-epochs 0
```

## Predict One Image

```powershell
python predict.py "garbage_classification\plastic\plastic1.jpg"
```

You can change the image path to any image you want to classify.

## Run Web App

```powershell
python app.py
```

Then open `http://127.0.0.1:5000` in your browser, upload an image, and choose
either saved model from the `models/` folder.

If port `5000` is unavailable, start it on another port:

```powershell
$env:PORT = "7860"
python app.py
```

## Deploy Live

The app is ready for a Python web host such as Render. Push this project to
GitHub, create a Render Web Service from the repository, and use:

- Build command: `pip install -r requirements.txt`
- Start command: `gunicorn --bind 0.0.0.0:$PORT app:app`
- Runtime: Python 3.11

This repository also includes `render.yaml`, so Render can detect the service
settings automatically from the repo.
