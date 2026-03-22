rm(list = ls())  # Clear the current R workspace

# Load required packages
library(flexsurv)   # Parametric survival modeling
library(survival)   # Surv() object
library(dplyr)      # Data manipulation
library(purrr)      # Functional programming helpers like map()/imap()
library(tidyr)      # pivot_wider()
library(tibble)     # Tidy tibble helpers
library(clipr)      # Copy output to clipboard

# Read tab-delimited data from the clipboard.
# Expected columns include at least: time, event
rfs <- read.table("clipboard", header = TRUE, sep = "\t")

# Fit one example generalized gamma model to the data.
# This is not used later in the main workflow; it looks like an initial test fit.
fit <- flexsurvreg(
  Surv(time, event) ~ 1,   # No covariates: intercept-only survival model
  data = rfs,
  dist = "gengamma.orig"
)

# Copy the input data to a more generic name
dat <- rfs

# Define the set of flexsurv distributions to fit.
# The names on the left are the labels that will appear in output tables.
# The values on the right are the flexsurv distribution names passed to flexsurvreg().
dist_map <- c(
  exp            = "exp",
  weibull        = "weibull",
  weibullPH      = "weibullPH",
  gompertz       = "gompertz",
  gamma          = "gamma",
  llogis         = "llogis",
  lnorm          = "lnorm",
  gengamma       = "gengamma",
  gengamma_orig  = "gengamma.orig",
  genf           = "genf",
  genf_orig      = "genf.orig"
)

# Function to fit a list of flexsurv models for multiple distributions
fit_flexsurv_models <- function(dat, time_var, event_var, dists = dist_map) {
  
  # Build the survival formula dynamically, e.g. Surv(time, event) ~ 1
  surv_formula <- as.formula(
    paste0("Surv(", time_var, ", ", event_var, ") ~ 1")
  )
  
  # Fit one model for each distribution in dists
  fits <- imap(
    dists,
    ~ flexsurvreg(
      formula = surv_formula,
      data    = dat,
      dist    = .x
    )
  )
  
  return(fits)
}

# Function to compute RMST (restricted mean survival time) up to time tau
# using numerical integration of the survival curve
compute_rmst_one <- function(fit_obj, tau, n_grid = 2000) {
  
  # Invalid tau values return NA
  if (is.na(tau) || tau < 0) {
    return(NA_real_)
  }
  
  # RMST from 0 to 0 is 0
  if (tau == 0) {
    return(0)
  }
  
  # Create a fine grid from 0 to tau for numerical integration
  grid <- seq(0, tau, length.out = n_grid)
  
  # Predict survival at each grid time
  sdat <- summary(fit_obj, t = grid, type = "survival")[[1]] %>%
    as_tibble() %>%
    select(time, est)
  
  tt <- sdat$time
  ss <- sdat$est
  
  # If there are too few points, integration is not possible
  if (length(tt) < 2 || length(ss) < 2) {
    return(NA_real_)
  }
  
  # Clamp survival values to [0, 1] just in case
  ss <- pmin(pmax(ss, 0), 1)
  
  # Trapezoidal numerical integration:
  # RMST = integral of S(t) dt from 0 to tau
  sum(diff(tt) * (head(ss, -1) + tail(ss, -1)) / 2)
}

# General function to predict different metrics from a list of fitted flexsurv models
predict_flexsurv_metric <- function(fit_list,
                                    times_vec = NULL,
                                    metric = "survival",
                                    probs_vec = c(0.25, 0.50, 0.75),
                                    rmst_grid_n = 2000) {
  
  # Restrict metric to supported values
  metric <- match.arg(metric, c("survival", "hazard", "cumhaz", "quantile", "rmst"))
  
  # Case 1: metrics evaluated at specified times
  if (metric %in% c("survival", "hazard", "cumhaz")) {
    
    if (is.null(times_vec)) {
      stop("times_vec must be provided for survival, hazard, and cumhaz")
    }
    
    # For each fitted model, get predictions at the supplied times
    out <- map_dfr(names(fit_list), function(dist_name) {
      
      fit_obj <- fit_list[[dist_name]]
      
      summary(fit_obj, t = times_vec, type = metric)[[1]] %>%
        as_tibble() %>%
        transmute(
          time  = time,
          dist  = dist_name,
          value = est
        )
    }) %>%
      # Convert from long to wide format:
      # one row per time, one column per distribution
      pivot_wider(
        names_from  = dist,
        values_from = value
      ) %>%
      arrange(time)
    
    names(out)[1] <- "time"
    return(out)
  }
  
  # Case 2: RMST evaluated at a vector of tau values
  if (metric == "rmst") {
    
    if (is.null(times_vec)) {
      stop("times_vec must be provided for rmst")
    }
    
    out <- map_dfr(names(fit_list), function(dist_name) {
      
      fit_obj <- fit_list[[dist_name]]
      
      tibble(
        tau   = times_vec,
        dist  = dist_name,
        value = map_dbl(times_vec, ~ compute_rmst_one(fit_obj, .x, n_grid = rmst_grid_n))
      )
    }) %>%
      # Convert from long to wide format:
      # one row per tau, one column per distribution
      pivot_wider(
        names_from  = dist,
        values_from = value
      ) %>%
      arrange(tau)
    
    return(out)
  }
  
  # Case 3: quantiles for each model
  if (metric == "quantile") {
    
    out <- map_dfr(names(fit_list), function(dist_name) {
      
      fit_obj <- fit_list[[dist_name]]
      
      summary(fit_obj, type = "quantile", quantiles = probs_vec)[[1]] %>%
        as_tibble() %>%
        transmute(
          prob  = quantile,
          dist  = dist_name,
          value = est
        )
    }) %>%
      # Convert from long to wide format:
      # one row per probability, one column per distribution
      pivot_wider(
        names_from  = dist,
        values_from = value
      ) %>%
      arrange(prob)
    
    # Add a readable quartile/percentile label
    out <- out %>%
      mutate(
        quartile = case_when(
          prob == 0.25 ~ "Q1",
          prob == 0.50 ~ "Median",
          prob == 0.75 ~ "Q3",
          TRUE         ~ paste0(prob * 100, "%")
        )
      ) %>%
      select(quartile, prob, everything())
    
    return(out)
  }
}

# Fit all requested flexsurv models to the dataset
fits <- fit_flexsurv_models(
  dat       = dat,
  time_var  = "time",
  event_var = "event"
)

# Predict hazard at times 0, 6, 12, ..., 240
haz_tbl <- predict_flexsurv_metric(
  fit_list  = fits,
  times_vec = seq(0, 240, 6),
  metric    = "hazard"
)

# Predict survival at times 0, 6, 12, ..., 240
surv_tbl <- predict_flexsurv_metric(
  fit_list  = fits,
  times_vec = seq(0, 240, 6),
  metric    = "survival"
)

# Predict cumulative hazard at times 0, 6, 12, ..., 240
cumhaz_tbl <- predict_flexsurv_metric(
  fit_list  = fits,
  times_vec = seq(0, 240, 6),
  metric    = "cumhaz"
)

# Compute RMST at tau = 12, 24, 36, ..., 240
rmst_tbl <- predict_flexsurv_metric(
  fit_list  = fits,
  times_vec = seq(0, 60, 12),
  metric    = "rmst"
)

# Compute quantiles at 10%, 20%, ..., 90%
quartile_tbl <- predict_flexsurv_metric(
  fit_list  = fits,
  metric    = "quantile",
  probs_vec = seq(0.1, 0.9, 0.1)
)

# Print output tables
haz_tbl
surv_tbl
cumhaz_tbl
rmst_tbl
quartile_tbl

# Copy quantile table to clipboard
clipr::write_clip(quartile_tbl)

# Optional: copy other tables to clipboard instead
# clipr::write_clip(rmst_tbl)
# clipr::write_clip(surv_tbl)
# clipr::write_clip(haz_tbl)
# clipr::write_clip(cumhaz_tbl)