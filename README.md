# flexsurv VBA Risk Functions

VBA implementations of risk functions for `flexsurv` and `flexsurvcure` parametric survival models.

## Supported distributions

- Exponential
- Weibull
- Gompertz
- Lognormal
- Loglogistic
- Generalized gamma
- Generalized F

## Functions included

- Survival (S)
- Density (PDF)
- Hazard (h)
- Cumulative hazard (H)
- Quantile functions

## Purpose

- Bring flexsurv-style survival modeling into Excel/VBA
- Enable reproducible parametric survival analysis outside R

## Example

```vba
Sweibull(time,shape,scale)


