# Capture NGS Variant Confidence Logistic Regression

Reproduction project for the paper:

> A machine learning model to determine the accuracy of variant calls in capture based next generation sequencing

The goal is a conservative binary classifier for variant calls:

- `1 = Present`: true variant
- `0 = Not present`: false/artifact variant

The model is intentionally calibrated to avoid false positives in the High confidence group.

## Contents

- `src/train_conservative_threshold.py`: end-to-end training and evaluation script.
- `data/raw/12864_2018_4659_MOESM2_ESM.xlsx`: Supplementary Table S2 variant-level data.
- `reports/variant-confidence-hust-report.pptx`: HUST report slide deck.
- `reports/final-contact-sheet.png`: rendered preview of the slide deck.
- `slides/build_report_deck.mjs`: script used to rebuild the slide deck from the HUST template workflow.

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

The script:

1. Reads the Excel supplementary table.
2. Builds the 14 feature columns:
   `DP, AD, AF, gc_5, gc_20, gc_50, MQ, GQ, WHR, HPL-D, HPL-L, QUAL, QD, FS`.
3. Encodes `Present -> 1`, `Not present -> 0`.
4. Splits data into Train/Dev/Test using `70/15/15` with stratification.
5. Trains `StandardScaler + LogisticRegression`.
6. Finds the lowest conservative threshold on Dev with `FP = 0`.
7. Evaluates once on Test and runs 10-fold CV AUC on the full dataset.

## Local Result

Current local run:

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

## Why Local Metrics Can Look Better Than The Published Result

The local result is not necessarily a true model improvement over the paper. It can look better for several methodological reasons:

- Different random split: this repo uses a new stratified 70/15/15 split, not necessarily the same held-out/test split used by the authors.
- Smaller test set: the local Test set has 1,077 variants, so a difference of a few variants changes the percentage noticeably.
- Different implementation details: scikit-learn Logistic Regression, solver behavior, regularization defaults, and preprocessing may not exactly match the original implementation.
- Threshold calibration differs: this repo chooses the lowest Dev threshold that gives `FP = 0`; the paper reports the conservative behavior but does not publish the exact threshold.
- Sampling variance: the positive/negative composition of Dev and Test can make AUC/FN/accuracy slightly better or worse even when the underlying model is effectively similar.
- Possible leakage is controlled, but not identical to the paper: the local pipeline keeps Test separate, but it still uses a split generated locally from the same supplementary dataset.

The safest conclusion is: the reproduction matches the paper's main behavior closely, especially `FP = 0` and High confidence around 92%, but it should not be claimed as a definitive improvement without the original split or external validation.

## Paper Links

- PubMed: https://pubmed.ncbi.nlm.nih.gov/29665779/
- PMC full text: https://pmc.ncbi.nlm.nih.gov/articles/PMC5904977/
