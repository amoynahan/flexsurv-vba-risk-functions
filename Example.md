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

## Survival Function Implementations

Parameters are taken directly from `flexsurv` output and mapped to VBA inputs.

Each function evaluates S(t) (survival probability) over the time grid `A18:A58`.

| Distribution      | VBA Function Call                      | Parameters                                                          | Notes                         |
| ----------------- | -------------------------------------- | ------------------------------------------------------------------- | ----------------------------- |
| Exponential       | `=sexp(A18:A58,B2)`                    | time = `A18:A58`, rate = `B2`                                       | Constant hazard               |
| Weibull           | `=Sweibull(A18:A58,B3,C3)`             | time = `A18:A58`, scale = `B3`, shape = `C3`                        | Matches flexsurv default      |
| WeibullPH         | `=SweibullPH(A18:A58,B4,C4)`           | time = `A18:A58`, scale = `B4`, shape = `C4`                        | Proportional hazards form     |
| Gompertz          | `=Sgompertz(A18:A58,B5,C5)`            | time = `A18:A58`, rate = `B5`, shape = `C5`                         | Exponentially changing hazard |
| Gamma             | `=Sgamma(A18:A58,B6,C6)`               | time = `A18:A58`, rate = `B6`, scale = `C6`                         | Flexible parametric form      |
| Log-logistic      | `=Sllogis(A18:A58,B7,C7)`              | time = `A18:A58`, shape = `B7`, scale = `C7`                        | Non-monotonic hazard          |
| Log-normal        | `=Slnorm(A18:A58,B8,C8)`               | time = `A18:A58`, meanlog = `B8`, sdlog = `C8`                      | Non-monotonic hazard          |
| Generalized Gamma | `=Sgengamma(A18:A58,B9,C9,D9)`         | time = `A18:A58`, mu = `B9`, sigma = `C9`, Q = `D9`                 | Includes Weibull/log-normal   |
| Generalized F     | `=Sgenf(A18:A58,B10,C10,D10,E10)`      | time = `A18:A58`, mu = `B10`, sigma = `C10`, Q = `D10`, P = `E10`   | Most flexible                 |
| Gen. Gamma (orig) | `=Sgengamma_orig(A18:A58,B11,C11,D11)` | time = `A18:A58`, shape = `B11`, scale = `C11`, k = `D11`           | flexsurv internal form        |
| Gen. F (orig)     | `=Sgenf_orig(A18:A58,B12,C12,D12,E12)` | time = `A18:A58`, mu = `B12`, sigma = `C12`, s1 = `D12`, s2 = `E12` | flexsurv internal form        |


using the same time grid (`A18:A58`) and distribution-specific parameters.
