# Model Card

## Model

Logistic Regression with a `StandardScaler` preprocessing step.

## Intended Use

Prioritize conservative High confidence variant calls for capture-based next generation sequencing data.

## Input Features

`DP, AD, AF, gc_5, gc_20, gc_50, MQ, GQ, WHR, HPL-D, HPL-L, QUAL, QD, FS`

## Target

- `1`: Present
- `0`: Not present

## Decision Policy

The model outputs `predict_proba(X)[:, 1]`. A variant is High confidence only if:

```text
probability >= conservative_threshold
```

The threshold is selected on the Development set as the lowest threshold with zero false positives.

## Limitations

- Results are based on a local Train/Dev/Test split and may change with a different random seed.
- The threshold should be recalibrated before use on a new panel, sequencer, pipeline, or lab dataset.
- This model should support variant review, not replace clinical validation.
