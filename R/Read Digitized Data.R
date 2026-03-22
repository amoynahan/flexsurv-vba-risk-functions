rm(list=ls())  # Clear all objects from the environment

# Load required libraries
library(clipr)        # For reading/writing clipboard data
library(IPDfromKM)    # For reconstructing IPD from Kaplan-Meier curves

############################################################
### Pembrolizumab Arm
############################################################

# Read digitized KM curve data from clipboard
# Expected columns: time, surv
PEMBRO_digdat <- read_clip_tbl()

# Read numbers-at-risk table from clipboard
# Expected columns: time, n.atrisk
PEMBRO_AtRisk <- read_clip_tbl()

# Total number of events reported in the study
total_events <- 103

#----------------------------------------------------------
# Preprocessing KM data
#----------------------------------------------------------

# Clamp survival probabilities to [0,1] to avoid invalid values
PEMBRO_digdat$surv <- pmin(pmax(PEMBRO_digdat$surv, 0), 1)

# Enforce monotonic decreasing survival (KM curves should not increase)
PEMBRO_digdat$surv <- cummin(PEMBRO_digdat$surv)

# Run preprocessing step required by IPDfromKM
# Aligns digitized KM with at-risk table and prepares intervals
PEMBRO_PP <- preprocess(
  PEMBRO_digdat,
  trisk = PEMBRO_AtRisk$time,
  nrisk = PEMBRO_AtRisk$n.atrisk,
  maxy  = 1.0
)

# Inspect processed data (optional)
PEMBRO_PP$preprocessdat

#----------------------------------------------------------
# Reconstruct individual patient data (IPD)
#----------------------------------------------------------

ipd <- getIPD(
  prep        = PEMBRO_PP,
  armID       = 0,
  tot.events  = total_events
)

# Extract reconstructed dataset
PEMBRO_IPD <- ipd$IPD

# Add unique patient ID
PEMBRO_IPD$id <- 1:nrow(PEMBRO_IPD)

# Add treatment label and rename status → event
PEMBRO_IPD <- PEMBRO_IPD %>%
  mutate(
    treatment = "Pembrolizumab",
    event     = status
  ) %>%
  select(id, treatment, time, event)

# Copy result to clipboard
clipr::write_clip(PEMBRO_IPD)

############################################################
### Chemotherapy Arm
############################################################

# Read digitized KM data
CHEMO_digdat <- read_clip_tbl()

# Read at-risk table
CHEMO_AtRisk <- read_clip_tbl()

# Total number of events
total_events <- 123

#----------------------------------------------------------
# Preprocessing KM data
#----------------------------------------------------------

# Clamp survival to [0,1]
CHEMO_digdat$surv <- pmin(pmax(CHEMO_digdat$surv, 0), 1)

# Enforce monotonic decreasing survival
CHEMO_digdat$surv <- cummin(CHEMO_digdat$surv)

# Preprocess KM + at-risk data
CHEMO_PP <- preprocess(
  CHEMO_digdat,
  trisk = CHEMO_AtRisk$time,
  nrisk = CHEMO_AtRisk$n.atrisk,
  maxy  = 1.0
)

#----------------------------------------------------------
# Reconstruct IPD
#----------------------------------------------------------

ipd <- getIPD(
  prep        = CHEMO_PP,
  armID       = 0,
  tot.events  = total_events
)

CHEMO_IPD <- ipd$IPD

# Add patient ID
CHEMO_IPD$id <- 1:nrow(CHEMO_IPD)

# Add treatment label and standardize columns
CHEMO_IPD <- CHEMO_IPD %>%
  mutate(
    treatment = "Chemotherapy",
    event     = status
  ) %>%
  select(id, treatment, time, event)

# Copy to clipboard
clipr::write_clip(CHEMO_IPD)

############################################################
### Combine Arms
############################################################

# Stack both treatment arms into a single dataset
os <- rbind(PEMBRO_IPD, CHEMO_IPD)

# Load survival analysis libraries
library(survival)
library(survminer)

# Inspect combined dataset
os

# Copy combined dataset to clipboard
clipr::write_clip(os)

############################################################
### Kaplan-Meier Analysis
############################################################

# Fit KM curves by treatment group
fit <- survfit(Surv(time, event) ~ treatment, data = os)

# Base R plot
plot(fit)

# Enhanced KM plot with risk table
ggsurvplot(
  fit,
  data          = os,
  conf.int      = FALSE,   # No confidence intervals
  risk.table    = TRUE,    # Show number at risk
  break.time.by = 6,       # X-axis ticks every 6 units
  xlim          = c(0, 72),# Limit x-axis
  ggtheme       = theme_bw()
)