"""Train logistic regression and find a conservative Dev threshold.

Default input:
    ../data/raw/12864_2018_4659_MOESM2_ESM.xlsx

This self-contained script reads the variant-level Excel file, builds X from
the 14 selected quality features, encodes y from confirmation, splits 70/15/15 with
stratification, trains StandardScaler + LogisticRegression, then finds the
lowest threshold that gives FP = 0 on the Dev split.
"""

from __future__ import annotations

import argparse
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

import numpy as np
import pandas as pd
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import confusion_matrix, roc_auc_score
from sklearn.model_selection import StratifiedKFold, cross_val_score, train_test_split
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler


DEFAULT_DATA_PATH = Path(__file__).resolve().parents[1] / "data" / "raw" / "12864_2018_4659_MOESM2_ESM.xlsx"
TARGET_COLUMN = "confirmation"
FEATURE_COLUMNS = [
    "DP",
    "AD",
    "AF",
    "gc_5",
    "gc_20",
    "gc_50",
    "MQ",
    "GQ",
    "WHR",
    "HPL-D",
    "HPL-L",
    "QUAL",
    "QD",
    "FS",
]


def col_index_from_ref(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    idx = 0
    for ch in letters:
        idx = idx * 26 + ord(ch.upper()) - ord("A") + 1
    return idx - 1


def xml_text(element: ElementTree.Element | None) -> str:
    if element is None:
        return ""
    return "".join(element.itertext())


def read_xlsx_without_openpyxl(path: Path) -> pd.DataFrame:
    ns = {
        "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
    }

    with zipfile.ZipFile(path) as archive:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ElementTree.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall("main:si", ns):
                shared_strings.append(xml_text(item))

        workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        rels = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_targets = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels.findall("pkgrel:Relationship", ns)
        }

        first_sheet = workbook.find("main:sheets/main:sheet", ns)
        if first_sheet is None:
            raise ValueError(f"No worksheets found in {path}")

        rel_id = first_sheet.attrib[f"{{{ns['rel']}}}id"]
        sheet_target = rel_targets[rel_id].lstrip("/")
        if not sheet_target.startswith("xl/"):
            sheet_target = f"xl/{sheet_target}"

        sheet_root = ElementTree.fromstring(archive.read(sheet_target))
        rows: list[list[str]] = []
        for row in sheet_root.findall(".//main:sheetData/main:row", ns):
            values: list[str] = []
            for cell in row.findall("main:c", ns):
                ref = cell.attrib.get("r", "")
                col_idx = col_index_from_ref(ref) if ref else len(values)
                while len(values) <= col_idx:
                    values.append("")

                cell_type = cell.attrib.get("t")
                if cell_type == "inlineStr":
                    value = xml_text(cell.find("main:is", ns))
                else:
                    raw_value = xml_text(cell.find("main:v", ns))
                    if cell_type == "s" and raw_value:
                        value = shared_strings[int(raw_value)]
                    elif cell_type == "b":
                        value = "TRUE" if raw_value == "1" else "FALSE"
                    else:
                        value = raw_value
                values[col_idx] = value
            if any(v != "" for v in values):
                rows.append(values)

    if not rows:
        raise ValueError(f"No rows found in {path}")

    header = [str(value).strip() for value in rows[0]]
    data_rows = [pad_row(row, len(header)) for row in rows[1:] if any(v != "" for v in row)]
    return pd.DataFrame(data_rows, columns=header)


def pad_row(row: list[str], width: int) -> list[str]:
    if len(row) >= width:
        return row[:width]
    return row + [""] * (width - len(row))


def load_input_table(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        return read_xlsx_without_openpyxl(path)
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path)
    raise ValueError(f"Unsupported input file type: {path.suffix}")


def build_xy(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.Series]:
    missing_columns = [col for col in [TARGET_COLUMN, *FEATURE_COLUMNS] if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Missing required columns: {missing_columns}")

    df = df.dropna(how="all").copy()

    y = (
        df[TARGET_COLUMN]
        .astype(str)
        .str.strip()
        .str.lower()
        .map({"present": 1, "not_present": 0})
    )
    if y.isna().any():
        bad_values = sorted(df.loc[y.isna(), TARGET_COLUMN].astype(str).unique())
        raise ValueError(f"Cannot encode target values: {bad_values}")

    x = df[FEATURE_COLUMNS].copy()
    for col in x.columns:
        x[col] = (
            x[col]
            .astype(str)
            .str.strip()
            .str.replace(",", ".", regex=False)
            .pipe(pd.to_numeric, errors="coerce")
        )

    if x.isna().any().any():
        missing_report = pd.DataFrame(
            {
                "missing_count": x.isna().sum(),
                "missing_rate": x.isna().mean(),
            }
        )
        raise ValueError(f"Feature matrix contains missing values:\n{missing_report}")

    return x, y.astype(int)


def split_70_15_15(
    x: pd.DataFrame,
    y: pd.Series,
    random_state: int,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.Series, pd.Series, pd.Series]:
    x_train, x_temp, y_train, y_temp = train_test_split(
        x,
        y,
        test_size=0.30,
        stratify=y,
        random_state=random_state,
    )
    x_dev, x_test, y_dev, y_test = train_test_split(
        x_temp,
        y_temp,
        test_size=0.50,
        stratify=y_temp,
        random_state=random_state,
    )
    return x_train, x_dev, x_test, y_train, y_dev, y_test


def find_conservative_threshold(y_true: pd.Series | np.ndarray, prob_present: np.ndarray) -> float:
    y_true = np.asarray(y_true).astype(int)
    prob_present = np.asarray(prob_present, dtype=float)

    negative_probs = prob_present[y_true == 0]
    if negative_probs.size == 0:
        return 0.0

    max_negative_prob = np.max(negative_probs)
    if max_negative_prob >= 1.0:
        raise ValueError("Cannot guarantee FP = 0 with threshold <= 1.0 on Dev set.")

    return float(np.nextafter(max_negative_prob, 1.0))


def make_pipeline(random_state: int) -> Pipeline:
    return Pipeline(
        steps=[
            ("scaler", StandardScaler()),
            (
                "logreg",
                LogisticRegression(
                    solver="lbfgs",
                    max_iter=10000,
                    random_state=random_state,
                ),
            ),
        ]
    )


def positive_class_probability(model: Pipeline, x: pd.DataFrame) -> np.ndarray:
    classes = model.named_steps["logreg"].classes_
    positive_class_index = np.where(classes == 1)[0][0]
    return model.predict_proba(x)[:, positive_class_index]


def evaluate_test_set(
    model: Pipeline,
    x_test: pd.DataFrame,
    y_test: pd.Series,
    conservative_threshold: float,
) -> dict[str, float | int | list[list[int]]]:
    y_true = np.asarray(y_test).astype(int)
    test_prob_present = positive_class_probability(model, x_test)
    test_pred_high_conf = (test_prob_present >= conservative_threshold).astype(int)

    tn, fp, fn, tp = confusion_matrix(y_true, test_pred_high_conf, labels=[0, 1]).ravel()

    return {
        "roc_auc": float(roc_auc_score(y_true, test_prob_present)),
        "confusion_matrix": [[int(tn), int(fp)], [int(fn), int(tp)]],
        "high_confidence_count": int(np.sum(test_pred_high_conf == 1)),
        "false_positive_count": int(fp),
        "present_low_confidence_count": int(fn),
    }


def cross_validate_auc(
    x: pd.DataFrame,
    y: pd.Series,
    random_state: int,
    n_splits: int = 10,
) -> np.ndarray:
    cv = StratifiedKFold(n_splits=n_splits, shuffle=True, random_state=random_state)
    return cross_val_score(
        estimator=make_pipeline(random_state),
        X=x,
        y=y,
        scoring="roc_auc",
        cv=cv,
        n_jobs=None,
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", type=Path, default=DEFAULT_DATA_PATH)
    parser.add_argument("--random-state", type=int, default=42)
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    df = load_input_table(args.data)
    x, y = build_xy(df)
    x_train, x_dev, x_test, y_train, y_dev, y_test = split_70_15_15(
        x=x,
        y=y,
        random_state=args.random_state,
    )

    pipeline = make_pipeline(args.random_state)
    pipeline.fit(x_train, y_train)

    dev_prob_present = positive_class_probability(pipeline, x_dev)
    conservative_threshold = find_conservative_threshold(
        y_true=y_dev,
        prob_present=dev_prob_present,
    )

    dev_pred_high_conf = (dev_prob_present >= conservative_threshold).astype(int)
    fp_count = int(np.sum((dev_pred_high_conf == 1) & (np.asarray(y_dev).astype(int) == 0)))
    high_conf_count = int(np.sum(dev_pred_high_conf == 1))

    assert fp_count == 0

    test_metrics = evaluate_test_set(
        model=pipeline,
        x_test=x_test,
        y_test=y_test,
        conservative_threshold=conservative_threshold,
    )
    cv_auc_scores = cross_validate_auc(
        x=x,
        y=y,
        random_state=args.random_state,
        n_splits=10,
    )

    print(f"data_path = {args.data}")
    print(f"train_shape = {x_train.shape}, dev_shape = {x_dev.shape}, test_shape = {x_test.shape}")
    print()
    print("Development threshold calibration")
    print(f"dev_false_positive_count = {fp_count}")
    print(f"dev_high_confidence_count = {high_conf_count}")
    print(f"conservative_threshold = {conservative_threshold:.10f}")
    print()
    print("Test set evaluation")
    print(f"test_roc_auc = {test_metrics['roc_auc']:.10f}")
    print("confusion_matrix [[TN, FP], [FN, TP]] =")
    print(test_metrics["confusion_matrix"])
    print(f"test_high_confidence_count = {test_metrics['high_confidence_count']}")
    print(f"test_false_positive_count = {test_metrics['false_positive_count']}")
    print(f"test_present_low_confidence_count = {test_metrics['present_low_confidence_count']}")
    print()
    print("10-fold cross-validation AUC on full dataset")
    print(f"cv_auc_scores = {[round(score, 10) for score in cv_auc_scores.tolist()]}")
    print(f"cv_auc_mean = {np.mean(cv_auc_scores):.10f}")
    print(f"cv_auc_std = {np.std(cv_auc_scores, ddof=1):.10f}")


if __name__ == "__main__":
    main()
