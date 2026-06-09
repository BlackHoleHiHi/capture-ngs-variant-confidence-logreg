# Capture NGS Variant Confidence Classifier

This is my personal machine learning project for classifying capture-based next generation sequencing variant calls into conservative confidence groups.

The main idea is simple: train a Logistic Regression model on variant-level quality features, then choose a conservative decision threshold so that false positives are avoided as much as possible.

## Project Goal

The model predicts whether a variant call is likely to be real:

- `1 = Present`: likely true variant
- `0 = Not present`: likely false or artifact variant

Instead of using the default `0.5` probability cutoff, this project searches for a conservative threshold on a Development set. A variant is called `High confidence` only when its predicted probability is above that threshold.

## Repository Structure

- `src/train_conservative_threshold.py`: training, threshold selection, test evaluation, and cross-validation.
- `data/raw/12864_2018_4659_MOESM2_ESM.xlsx`: public variant-level Excel dataset used for the experiment.
- `reports/variant-confidence-hust-report.pptx`: presentation slide deck for reporting the project.
- `reports/final-contact-sheet.png`: rendered preview of the slide deck.
- `slides/build_report_deck.mjs`: script used to rebuild the slide deck.
- `MODEL_CARD.md`: short summary of model behavior, target, and limitations.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Run

```powershell
python .\src\train_conservative_threshold.py
```

## Pipeline

1. Read the Excel dataset.
2. Clean empty or metadata-like rows.
3. Encode the target:
   - `Present -> 1`
   - `Not present -> 0`
4. Build 14 numerical features:
   `DP, AD, AF, gc_5, gc_20, gc_50, MQ, GQ, WHR, HPL-D, HPL-L, QUAL, QD, FS`.
5. Split the dataset into Train/Dev/Test with a `70/15/15` ratio.
6. Train a `StandardScaler + LogisticRegression` pipeline.
7. Select the lowest Dev threshold that produces `FP = 0`.
8. Evaluate the final model on the held-out Test set.
9. Run 10-fold cross-validation AUC as an additional sanity check.

## Current Result

```text
conservative_threshold = 0.4730547338
test_roc_auc = 0.9998311688
confusion_matrix [[TN, FP], [FN, TP]] = [[77, 0], [4, 996]]
test_high_confidence_count = 996
test_false_positive_count = 0
test_present_low_confidence_count = 4
cv_auc_mean = 0.9997757219
cv_auc_std = 0.0002844816
```

## Interpretation

The current split gives a very strong result: the model keeps false positives at zero on the Test set while still calling most true variants as High confidence.

This should be interpreted carefully. A strong result on one split does not prove that the model is ready for clinical use. The threshold and performance should be checked again on new data from the same sequencing panel, pipeline, and lab workflow.

## Why The Metrics Can Look Very High

Several factors can make the local metrics look especially strong:

- The dataset is highly separable with the selected quality features.
- The Test set is relatively small, so a few samples can noticeably change percentages.
- The threshold is tuned on the Dev set for a conservative operating point.
- The Train/Dev/Test split is generated locally, so performance may change with a different random seed.
- Cross-validation AUC is a ranking metric and does not directly measure the conservative threshold behavior.

The safest conclusion is that this project demonstrates a useful conservative variant-confidence workflow, but it still needs external validation before being used in a real diagnostic setting.
