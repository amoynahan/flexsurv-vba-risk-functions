# Example: VBA Survival Functions (flexsurv Alignment)

## Overview

This example is based on the **"Survival-Test"** tab from the Excel workbook:

**Flexsurv functions with flexsurv cure functions.xlsm**

The sheet demonstrates how survival functions implemented in Excel VBA reproduce results from `flexsurv` in R.

Model parameters estimated in R are mapped to distribution-specific inputs and then evaluated using VBA survival functions to compute S(t) across a time grid.

---

## Example Workbook (Survival-Test Tab)

![Survival Workbook](docs/images/SurvivalFunctions.jpg)

---

## What the Sheet Is Doing

The workbook serves as a validation bridge:

- Parameters are estimated in R (`flexsurv`)  
- Parameters are mapped to the correct inputs for each distribution  
- VBA functions compute survival probabilities  
- Results can be compared directly to R outputs  

Each column represents a distribution, and each row represents a time point.

---

## Survival Function Implementations

Each formula evaluates:

S(t) = survival probability at time t

using the same time grid (`A18:A58`) and distribution-specific parameters.

S(t) = survival probability at time t

using the same time grid (`A18:A58`) and distribution-specific parameters.
