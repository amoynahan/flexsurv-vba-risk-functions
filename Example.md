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

Each function evaluates S(t) (survival probability) over the time grid `A18:A58` using parameters from the workbook.

| Distribution | VBA Function Call | Parameters | Notes |
|-------------|------------------|-----------|------|
| Exponential | `=sexp(A18:A58, B2)` | rate | Constant hazard |
| Weibull | `=Sweibull(A18:A58, B3, C3)` | scale, shape | Matches flexsurv default |
| WeibullPH | `=SweibullPH(A18:A58, B4, C4)` | scale, shape | Proportional hazards form |
| Gompertz | `=Sgompertz(A18:A58, B5, C5)` | rate, shape | Exponentially changing hazard |
| Gamma | `=Sgamma(A18:A58, B6, C6)` | rate, scale | Flexible parametric form |
| Log-logistic | `=Sllogis(A18:A58, B7, C7)` | shape, scale | Non-monotonic hazard |
| Log-normal | `=Slnorm(A18:A58, B8, C8)` | meanlog, sdlog | Non-monotonic hazard |
| Generalized Gamma | `=Sgengamma(A18:A58, B9, C9, D9)` | mu, sigma, Q | Includes Weibull/log-normal |
| Gen. Gamma (orig) | `=Sgengamma_orig(A18:A58, B11, C11, D11)` | mu, sigma, Q | flexsurv internal form |
| Generalized F | `=Sgenf(A18:A58, B10, C10, D10, E10)` | mu, sigma, Q, P | Most flexible |
| Gen. F (orig) | `=Sgenf_orig(A18:A58, B12, C12, D12, E12)` | mu, sigma, Q, P | flexsurv internal form |



using the same time grid (`A18:A58`) and distribution-specific parameters.
