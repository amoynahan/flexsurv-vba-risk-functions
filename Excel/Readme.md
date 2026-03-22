## Excel Workbook

### Overview

This workbook provides a working Excel/VBA environment for evaluating parametric survival models using risk functions aligned with the `flexsurv` and `flexsurvcure` frameworks.

It is intended as a companion to the VBA modules in this repository and includes example data, test sheets, and visualizations.

---

### Contents

The workbook includes:

- VBA-driven parametric survival risk functions  
- Example datasets for testing and validation  
- Calculation sheets for survival, hazard, and related functions  
- Plotting sheets for visualizing model outputs  
- Support for both standard survival models and mixture cure models  
- A Guyot algorithm implementation for reconstructing IPD from Kaplan–Meier curves  

---

### Workbook Structure

#### Core Survival Modeling

- **Data**  
  Input data for standard survival model testing  

- **Survival-Test**  
  Evaluation of survival, hazard, density, and related functions  

- **Survival-Plot**  
  Visualization of survival curves generated from VBA functions  

---

#### Flexible Parametric / Spline Models

- **RCS-test**  
  Testing of restricted cubic spline–based survival functions  

- **RCS-Plot**  
  Visualization of spline-based survival behavior  

---

#### Cure Models (`flexsurvcure`)

- **Data Cure**  
  Input data for mixture cure models  

- **Cure Test**  
  Evaluation of cure model survival and risk functions  

- **Cure Plot**  
  Visualization of cure model outputs  

- **Data Cure Estimate**  
  Supporting calculations for cure model estimation  

---

#### Kaplan–Meier Reconstruction (Supporting Utility)

- **Guyot Algorithm**  
  Implementation of the Guyot method for reconstructing individual patient data (IPD) from Kaplan–Meier curves  

---

### Purpose

This workbook is designed to:

- Provide a practical Excel interface for parametric survival modeling  
- Enable use of `flexsurv` and `flexsurvcure`-style risk functions in VBA  
- Support validation and visualization of survival distributions  
- Demonstrate how the VBA functions can be used in a spreadsheet environment  

---

### Notes

- The primary functionality is implemented in the VBA modules included in this repository  
- The Excel sheets are intended for testing, demonstration, and visualization  
- Kaplan–Meier reconstruction tools are included as supporting functionality, but are not the main focus  
