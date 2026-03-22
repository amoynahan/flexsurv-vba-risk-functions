## R Scripts Overview

This directory contains R scripts supporting survival analysis tasks, including parametric modeling, spline-based modeling, and mixture cure models. The scripts are modular and can be used independently depending on the analysis needs.

Below is a summary of the purpose of each file.

### FlexsurvModels.R

This script fits standard parametric survival models using the `flexsurv` package.

It includes:
- fitting multiple distributions (e.g., exponential, Weibull, Gompertz, log-normal, generalized gamma, generalized F)
- generating model-based predictions (survival, hazard, cumulative hazard, quantiles, RMST)
- extracting model coefficients into structured long and wide formats

This file is used for comparing parametric model specifications and producing summary outputs.

---

### RCS Models.R

This script fits flexible spline-based survival models using `flexsurvspline`.

It provides:
- spline-based models across different scales (hazard, odds, normal)
- estimation of survival, hazard, cumulative hazard, quantiles, and RMST
- numerical integration for RMST
- structured output tables and parameter extraction

This file is used when more flexible functional forms are desired beyond standard parametric models.

---

### Mixture Models.R

This script fits mixture cure models using the `flexsurvcure` package.

It supports:
- estimation of cure fractions alongside survival distributions
- fitting models across multiple baseline distributions
- use of different link functions for the cure component
- generation of outputs consistent with non-cure model summaries

This file is used when a cured fraction is considered relevant for the analysis.

---

### Read Digitized Data.R

This script reconstructs approximate individual patient data (IPD) from digitized Kaplan–Meier curves and numbers-at-risk tables using the `IPDfromKM` package.

It performs:
- input of digitized survival data and at-risk counts
- preprocessing to enforce valid survival curves
- reconstruction of event-time data
- creation of treatment-arm datasets and combined datasets
- basic Kaplan–Meier visualization

This file is used when patient-level data are not directly available and must be approximated from published results.

---

## Summary

These scripts provide a set of utilities for fitting and summarizing survival models under different assumptions and levels of flexibility. They can be used individually or in combination depending on the analysis requirements.
