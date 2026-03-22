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

## Distribution Details

Each formula evaluates:

S(t) = survival probability at time t

using parameters from the workbook.

---

=sexp(A18:A58, B2)


- Parameter: rate  
- Survival: S(t) = exp(-rate × t)  
- Constant hazard  

---

### Weibull


=Sweibull(A18:A58, B3, C3)


- Parameters: scale, shape  
- Flexible monotonic hazard  
- Matches `flexsurv` default parameterization  

---

### WeibullPH


=SweibullPH(A18:A58, B4, C4)


- Parameters: scale, shape  
- Proportional hazards parameterization  
- Same distribution, different interpretation  

---

### Gompertz


=Sgompertz(A18:A58, B5, C5)


- Parameters: rate, shape  
- Exponentially changing hazard  

---

### Gamma


=Sgamma(A18:A58, B6, C6)


- Parameters: rate, scale  
- Flexible but less commonly used  

---

### Log-logistic


=Sllogis(A18:A58, B7, C7)


- Parameters: shape, scale  
- Non-monotonic hazard  

---

### Log-normal


=Slnorm(A18:A58, B8, C8)


- Parameters: meanlog, sdlog  
- Non-monotonic hazard  

---

### Generalized Gamma


=Sgengamma(A18:A58, B9, C9, D9)


- Parameters: mu, sigma, Q  
- Highly flexible  
- Includes Weibull and log-normal as special cases  

---

### Generalized Gamma (Original)


=Sgengamma_orig(A18:A58, B11, C11, D11)


- Alternative parameterization used by `flexsurv`  
- Included for exact replication  

---

### Generalized F


=Sgenf(A18:A58, B10, C10, D10, E10)


- Parameters: mu, sigma, Q, P  
- Most flexible distribution in this set  

---

### Generalized F (Original)


=Sgenf_orig(A18:A58, B12, C12, D12, E12)


- Alternative parameterization used by `flexsurv`  
- Included for consistency with R outputs  

---

## Notes

- All formulas use the same time grid (`A18:A58`)  
- Parameters are taken directly from `flexsurv` outputs  
- Correct parameter mapping is critical  
- Survival at time 0 equals 1.0 for all distributions  

---


### Exponential
