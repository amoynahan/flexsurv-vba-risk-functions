# Example: VBA Survival Functions (flexsurv Alignment)

## Overview

This example is based on the **"Survival-Test"** tab from the Excel workbook:

**Flexsurv functions with flexsurv cure functions.xlsm**

The sheet demonstrates how survival functions implemented in Excel VBA reproduce results from `flexsurv` in R.

Model parameters estimated in R are mapped to distribution-specific inputs and then evaluated using VBA survival functions to compute S(t) across a time grid.

---

## Example Workbook

![Survival Workbook](docs/images/SurvivalFunctions.jpg)

---

## Distribution Details

Each column in the workbook corresponds to a distribution. The formulas evaluate:

S(t) = survival probability at time t

using parameters estimated in R.

---

### Exponential
