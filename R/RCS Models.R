rm(list=ls())  # Clear all objects from the environment

# Load required libraries
library(flexsurv)   # Flexible parametric survival models
library(survival)   # Core survival analysis functions (Surv, etc.)
library(dplyr)      # Data manipulation
library(purrr)      # Functional programming (map, imap)
library(tidyr)      # Data reshaping (pivot_wider)
library(tibble)     # Modern data frames
library(clipr)      # Clipboard interaction

# Read input data from clipboard (expects tab-delimited with header)
rfs <- read.table("clipboard", header = TRUE, sep = "\t")

# Copy dataset (for consistency / future flexibility)
dat <- rfs

# Define mapping of spline scales used in flexsurvspline
scale_map <- c(
  hazard = "hazard",   # Hazard scale spline
  odds   = "odds",     # Odds scale spline
  normal = "normal"    # Normal (log cumulative hazard) scale
)

#------------------------------------------------------------
# Function: Fit flexsurvspline models for each scale
#------------------------------------------------------------
fit_flexsurvspline_models <- function(dat,
                                      time_var,
                                      event_var,
                                      scales = scale_map,
                                      k = 1) {
  
  # Construct survival formula dynamically
  surv_formula <- as.formula(
    paste0("Surv(", time_var, ", ", event_var, ") ~ 1")
  )
  
  # Fit spline models for each scale (hazard, odds, normal)
  fits <- imap(
    scales,
    ~ flexsurvspline(
      formula = surv_formula,
      data    = dat,
      scale   = .x,   # scale type
      k       = k     # number of spline knots
    )
  )
  
  return(fits)  # Named list of fitted models
}

#------------------------------------------------------------
# Function: Compute RMST (Restricted Mean Survival Time)
#------------------------------------------------------------
compute_rmst_one_rcs <- function(fit_obj, tau, n_grid = 2000) {
  
  # Handle invalid inputs
  if (is.na(tau) || tau < 0) {
    return(NA_real_)
  }
  
  if (tau == 0) {
    return(0)
  }
  
  # Create time grid from 0 to tau
  grid <- seq(0, tau, length.out = n_grid)
  
  # Get survival estimates at grid points
  sdat <- summary(fit_obj, t = grid, type = "survival")[[1]] %>%
    as_tibble() %>%
    select(time, est)
  
  tt <- sdat$time  # time points
  ss <- sdat$est   # survival probabilities
  
  # Require at least 2 points for integration
  if (length(tt) < 2 || length(ss) < 2) {
    return(NA_real_)
  }
  
  # Clamp survival values to [0,1] for stability
  ss <- pmin(pmax(ss, 0), 1)
  
  # Trapezoidal integration of survival curve → RMST
  sum(diff(tt) * (head(ss, -1) + tail(ss, -1)) / 2)
}

#------------------------------------------------------------
# Function: Predict metrics from fitted models
#------------------------------------------------------------
predict_flexsurvspline_metric <- function(fit_list,
                                          times_vec = NULL,
                                          metric = "survival",
                                          probs_vec = c(0.25, 0.50, 0.75),
                                          rmst_grid_n = 2000) {
  
  # Validate metric input
  metric <- match.arg(metric, c("survival", "hazard", "cumhaz", "quantile", "rmst"))
  
  #----------------------------------------------------------
  # Case 1: Time-based outputs (survival, hazard, cumulative hazard)
  #----------------------------------------------------------
  if (metric %in% c("survival", "hazard", "cumhaz")) {
    
    if (is.null(times_vec)) {
      stop("times_vec must be provided for survival, hazard, and cumhaz")
    }
    
    # Loop over models and extract predictions
    out <- map_dfr(names(fit_list), function(scale_name) {
      
      fit_obj <- fit_list[[scale_name]]
      
      summary(fit_obj, t = times_vec, type = metric)[[1]] %>%
        as_tibble() %>%
        transmute(
          time  = time,
          scale = scale_name,
          value = est
        )
    }) %>%
      pivot_wider(               # reshape: one column per scale
        names_from  = scale,
        values_from = value
      ) %>%
      arrange(time)
    
    names(out)[1] <- "time"
    return(out)
  }
  
  #----------------------------------------------------------
  # Case 2: RMST (integration-based)
  #----------------------------------------------------------
  if (metric == "rmst") {
    
    if (is.null(times_vec)) {
      stop("times_vec must be provided for rmst")
    }
    
    out <- map_dfr(names(fit_list), function(scale_name) {
      
      fit_obj <- fit_list[[scale_name]]
      
      tibble(
        tau   = times_vec,
        scale = scale_name,
        value = map_dbl(times_vec, ~ compute_rmst_one_rcs(fit_obj, .x, n_grid = rmst_grid_n))
      )
    }) %>%
      pivot_wider(
        names_from  = scale,
        values_from = value
      ) %>%
      arrange(tau)
    
    return(out)
  }
  
  #----------------------------------------------------------
  # Case 3: Quantiles (e.g., median survival)
  #----------------------------------------------------------
  if (metric == "quantile") {
    
    out <- map_dfr(names(fit_list), function(scale_name) {
      
      fit_obj <- fit_list[[scale_name]]
      
      summary(fit_obj, type = "quantile", quantiles = probs_vec)[[1]] %>%
        as_tibble() %>%
        transmute(
          prob  = quantile,
          scale = scale_name,
          value = est
        )
    }) %>%
      pivot_wider(
        names_from  = scale,
        values_from = value
      ) %>%
      arrange(prob)
    
    # Add readable labels (Q1, Median, etc.)
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

#------------------------------------------------------------
# Function: Extract coefficients (long format)
#------------------------------------------------------------
extract_flexsurvspline_coefs_long <- function(fit_list) {
  
  map_dfr(names(fit_list), function(scale_name) {
    
    fit_obj <- fit_list[[scale_name]]
    res_df  <- as.data.frame(fit_obj$res)  # model results
    
    tibble(
      scale     = scale_name,
      param_raw = rownames(res_df),
      estimate  = as.numeric(res_df$est)
    ) %>%
      mutate(param = param_raw) %>%
      select(scale, param, estimate)
  })
}

#------------------------------------------------------------
# Function: Extract coefficients (wide format + parameter map)
#------------------------------------------------------------
extract_flexsurvspline_coefs_wide <- function(fit_list) {
  
  coef_long <- extract_flexsurvspline_coefs_long(fit_list)
  
  # Assign column indices per parameter within each scale
  coef_long2 <- coef_long %>%
    group_by(scale) %>%
    mutate(
      output_col = paste0("param", row_number())
    ) %>%
    ungroup()
  
  # Wide table: one row per model
  coef_wide <- coef_long2 %>%
    select(scale, output_col, estimate) %>%
    pivot_wider(
      names_from  = output_col,
      values_from = estimate
    ) %>%
    arrange(scale)
  
  # Parameter map: tells what each column represents
  param_map_wide <- coef_long2 %>%
    select(scale, output_col, param) %>%
    pivot_wider(
      names_from  = output_col,
      values_from = param
    ) %>%
    arrange(scale)
  
  list(
    coef_wide      = coef_wide,
    param_map_wide = param_map_wide,
    coef_long      = coef_long2
  )
}

#------------------------------------------------------------
# Fit models
#------------------------------------------------------------
fits_rcs <- fit_flexsurvspline_models(
  dat       = dat,
  time_var  = "time",
  event_var = "event",
  k         = 1
)

#------------------------------------------------------------
# Generate outputs
#------------------------------------------------------------
haz_tbl_rcs <- predict_flexsurvspline_metric(
  fit_list  = fits_rcs,
  times_vec = seq(0, 240, 6),
  metric    = "hazard"
)

surv_tbl_rcs <- predict_flexsurvspline_metric(
  fit_list  = fits_rcs,
  times_vec = seq(0, 240, 6),
  metric    = "survival"
)

cumhaz_tbl_rcs <- predict_flexsurvspline_metric(
  fit_list  = fits_rcs,
  times_vec = seq(0, 240, 6),
  metric    = "cumhaz"
)

rmst_tbl_rcs <- predict_flexsurvspline_metric(
  fit_list  = fits_rcs,
  times_vec = seq(0, 60, 12),
  metric    = "rmst"
)

quartile_tbl_rcs <- predict_flexsurvspline_metric(
  fit_list  = fits_rcs,
  metric    = "quantile",
  probs_vec = seq(0.1, 0.9, 0.1)
)

# Extract coefficients
coef_out_rcs <- extract_flexsurvspline_coefs_wide(fits_rcs)

coef_wide_rcs      <- coef_out_rcs$coef_wide
param_map_wide_rcs <- coef_out_rcs$param_map_wide
coef_long_rcs      <- coef_out_rcs$coef_long

#------------------------------------------------------------
# Output results
#------------------------------------------------------------
haz_tbl_rcs
surv_tbl_rcs
cumhaz_tbl_rcs
rmst_tbl_rcs
quartile_tbl_rcs
coef_wide_rcs
param_map_wide_rcs
coef_long_rcs

# Copy quartiles to clipboard
clipr::write_clip(quartile_tbl_rcs)

# Optional clipboard outputs
# clipr::write_clip(rmst_tbl_rcs)
# clipr::write_clip(surv_tbl_rcs)
# clipr::write_clip(haz_tbl_rcs)
# clipr::write_clip(cumhaz_tbl_rcs)
# clipr::write_clip(coef_wide_rcs)
# clipr::write_clip(param_map_wide_rcs)