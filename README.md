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
- RCS (odds, normal, hazard)
- Mixture models

## Functions included

- Survival (S)
- Density (d)
- Hazard (h)
- Cumulative hazard (ch)
- Quantile functions

## Purpose

- Bring flexsurv-style survival modeling into Excel/VBA
- Enable reproducible parametric survival analysis outside R

## Example

```vba
Sweibull(time,shape,scale)


