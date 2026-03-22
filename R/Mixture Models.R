rm(list = ls())  # Clear the current R workspace

# Load required packages
library(flexsurv)       # Parametric survival models
library(flexsurvcure)   # Mixture cure models built on flexsurv
library(survival)       # Surv() object
library(dplyr)          # Data manipulation
library(purrr)          # map(), imap(), map_dbl()
library(tidyr)          # pivot_wider()
library(tibble)         # Tibbles
library(clipr)          # Copy output to clipboard

# Read tab-delimited data from the clipboard
# Expected columns include at least: time, event
os <- read.table("clipboard", header = TRUE, sep = "\t")
dat <- os

# Define the set of baseline distributions to use inside flexsurvcure
# The names on the left become output labels
# The values on the right are the distribution names passed to flexsurvcure()
dist_map <- c(
  exp           = "exp",
  weibull       = "weibull",
  weibullPH     = "weibullPH",
  gompertz      = "gompertz",
  gamma         = "gamma",
  llogis        = "llogis",
  lnorm         = "lnorm",
  gengamma      = "gengamma",
  gengamma_orig = "gengamma.orig",
  genf          = "genf",
  genf_orig     = "genf.orig"
)

# Fit a set of flexsurvcure mixture cure models
fit_flexsurvcure_models <- function(dat,
                                    time_var,
                                    event_var,
                                    dists = dist_map,
                                    link = "logistic") {
  
  # Build survival formula dynamically, e.g. Surv(time, event) ~ 1
  surv_formula <- as.formula(
    paste0("Surv(", time_var, ", ", event_var, ") ~ 1")
  )
  
  # Fit one cure model per distribution
  # mixture = TRUE means a mixture cure model
  # link controls how the cure fraction is modeled
  # tryCatch prevents one failed model from stopping the whole loop
  fits <- imap(
    dists,
    ~ tryCatch(
      flexsurvcure(
        formula = surv_formula,
        data    = dat,
        dist    = .x,
        link    = link,
        mixture = TRUE
      ),
      error = function(e) {
        message("Model failed for ", .y, ": ", e$message)
        NULL
      }
    )
  )
  
  # Drop any failed fits (NULL entries)
  fits[!map_lgl(fits, is.null)]
}

# Compute RMST (restricted mean survival time) up to tau
# for one fitted flexsurvcure model using numerical integration
compute_rmst_one_cure <- function(fit_obj, tau, n_grid = 2000) {
  
  # Invalid tau values return NA
  if (is.na(tau) || tau < 0) {
    return(NA_real_)
  }
  
  # RMST from 0 to 0 is 0
  if (tau == 0) {
    return(0)
  }
  
  # Create a fine grid from 0 to tau
  grid <- seq(0, tau, length.out = n_grid)
  
  # Predict survival on the grid
  # tryCatch protects against summary() failures for some models
  sdat <- tryCatch(
    summary(fit_obj, t = grid, type = "survival")[[1]] %>%
      as_tibble() %>%
      select(time, est),
    error = function(e) NULL
  )
  
  # If prediction failed or returned too few points, return NA
  if (is.null(sdat) || nrow(sdat) < 2) {
    return(NA_real_)
  }
  
  tt <- sdat$time
  ss <- sdat$est
  
  # Clamp survival values to [0,1]
  ss <- pmin(pmax(ss, 0), 1)
  
  # Trapezoidal integration of the survival curve
  sum(diff(tt) * (head(ss, -1) + tail(ss, -1)) / 2)
}

# Predict different metrics from a list of fitted flexsurvcure models
predict_flexsurvcure_metric <- function(fit_list,
                                        times_vec = NULL,
                                        metric = c("survival", "hazard", "cumhaz", "quantile", "rmst"),
                                        probs_vec = c(0.25, 0.50, 0.75),
                                        rmst_grid_n = 2000) {
  
  metric <- match.arg(metric)
  
  # Case 1: metrics evaluated at specific times
  if (metric %in% c("survival", "hazard", "cumhaz")) {
    
    if (is.null(times_vec)) {
      stop("times_vec must be provided for survival, hazard, and cumhaz.")
    }
    
    # For each fitted model, get predictions at the supplied times
    out <- imap_dfr(fit_list, function(fit_obj, dist_name) {
      summary(fit_obj, t = times_vec, type = metric)[[1]] %>%
        as_tibble() %>%
        transmute(
          time  = time,
          dist  = dist_name,
          value = est
        )
    }) %>%
      # Reshape to one row per time, one column per distribution
      pivot_wider(
        names_from  = dist,
        values_from = value
      ) %>%
      arrange(time)
    
    return(out)
  }
  
  # Case 2: RMST at specified truncation times
  if (metric == "rmst") {
    
    if (is.null(times_vec)) {
      stop("times_vec must be provided for rmst. These are the truncation times (tau).")
    }
    
    out <- imap_dfr(fit_list, function(fit_obj, dist_name) {
      tibble(
        tau   = times_vec,
        dist  = dist_name,
        value = map_dbl(times_vec, ~ compute_rmst_one_cure(fit_obj, .x, n_grid = rmst_grid_n))
      )
    }) %>%
      # Reshape to one row per tau, one column per distribution
      pivot_wider(
        names_from  = dist,
        values_from = value
      ) %>%
      arrange(tau)
    
    return(out)
  }
  
  # Case 3: quantiles
  if (metric == "quantile") {
    
    out <- imap_dfr(fit_list, function(fit_obj, dist_name) {
      
      # Extract parameter results table from model object
      res_df <- as.data.frame(fit_obj$res)
      
      # Look for theta, the cure-model parameter associated with cure fraction
      theta_row <- grep("^theta$", rownames(res_df), ignore.case = TRUE)
      
      # If theta is not found, return NA for all requested probabilities
      if (length(theta_row) == 0) {
        warning("No theta found for ", dist_name)
        return(tibble(
          prob  = probs_vec,
          dist  = dist_name,
          value = NA_real_
        ))
      }
      
      # Estimated theta parameter
      theta_hat <- as.numeric(res_df$est[theta_row[1]])
      
      # In a mixture cure model, quantiles above (1 - theta_hat) may not exist
      # because the survival curve levels off above zero
      valid_probs <- probs_vec[probs_vec < (1 - theta_hat)]
      
      # Compute only the quantiles that are theoretically defined
      q_tbl <- tryCatch(
        {
          if (length(valid_probs) > 0) {
            summary(fit_obj, type = "quantile", quantiles = valid_probs)[[1]] %>%
              as_tibble() %>%
              mutate(prob = valid_probs[seq_len(n())]) %>%
              transmute(
                prob  = prob,
                dist  = dist_name,
                value = est
              )
          } else {
            tibble(
              prob  = numeric(0),
              dist  = character(0),
              value = numeric(0)
            )
          }
        },
        error = function(e) {
          warning("Quantile failed for ", dist_name, ": ", e$message)
          tibble(
            prob  = valid_probs,
            dist  = dist_name,
            value = NA_real_
          )
        }
      )
      
      # Re-expand back to the full requested probability vector,
      # filling undefined quantiles with NA
      tibble(prob = probs_vec) %>%
        left_join(q_tbl, by = "prob") %>%
        mutate(dist = dist_name) %>%
        select(prob, dist, value)
    }) %>%
      # Reshape to one row per probability, one column per distribution
      pivot_wider(
        names_from  = dist,
        values_from = value
      ) %>%
      arrange(prob)
    
    return(out)
  }
}

# Extract all coefficients from each flexsurvcure model into long format
extract_flexsurvcure_coefs_long <- function(fit_list) {
  
  imap_dfr(fit_list, function(fit_obj, dist_name) {
    
    res_df <- as.data.frame(fit_obj$res)
    
    tibble(
      dist      = dist_name,
      param_raw = rownames(res_df),
      estimate  = as.numeric(res_df$est)
    ) %>%
      mutate(
        # Standardize theta naming
        param = case_when(
          grepl("^theta$", param_raw, ignore.case = TRUE) ~ "theta",
          TRUE                                            ~ param_raw
        )
      ) %>%
      select(dist, param, estimate)
  })
}

# Convert coefficient output to wide format and create a parameter map
extract_flexsurvcure_coefs_wide <- function(fit_list) {
  
  coef_long <- extract_flexsurvcure_coefs_long(fit_list)
  
  coef_long2 <- coef_long %>%
    group_by(dist) %>%
    mutate(
      # Preserve parameter order within each distribution
      param_order = row_number(),
      # Force theta to appear first
      param_order = if_else(param == "theta", 0L, param_order)
    ) %>%
    arrange(param_order, .by_group = TRUE) %>%
    mutate(
      # Create generic output column names
      output_col = paste0("param", row_number())
    ) %>%
    ungroup()
  
  # Wide coefficient table: estimates only
  coef_wide <- coef_long2 %>%
    select(dist, output_col, estimate) %>%
    pivot_wider(
      names_from  = output_col,
      values_from = estimate
    ) %>%
    arrange(dist)
  
  # Wide map of which parameter is in each generic column
  param_map_wide <- coef_long2 %>%
    select(dist, output_col, param) %>%
    pivot_wider(
      names_from  = output_col,
      values_from = param
    ) %>%
    arrange(dist)
  
  list(
    coef_wide      = coef_wide,
    param_map_wide = param_map_wide,
    coef_long      = coef_long2
  )
}

# Fit all requested mixture cure models
fits <- fit_flexsurvcure_models(
  dat       = dat,
  time_var  = "time",
  event_var = "event"
)

# Predict hazard at 0, 6, 12, ..., 240
haz_tbl <- predict_flexsurvcure_metric(
  fit_list  = fits,
  times_vec = seq(0, 240, 6),
  metric    = "hazard"
)

# Predict survival at 0, 6, 12, ..., 240, plus long-tail times
# The extra 2400 and 3400 help inspect long-run cure-model behavior
surv_tbl <- predict_flexsurvcure_metric(
  fit_list  = fits,
  times_vec = c(seq(0, 240, 6), 2400, 3400),
  metric    = "survival"
)

# Predict cumulative hazard at 0, 6, 12, ..., 240
cumhaz_tbl <- predict_flexsurvcure_metric(
  fit_list  = fits,
  times_vec = seq(0, 240, 6),
  metric    = "cumhaz"
)

# Compute RMST at tau = 0, 12, ..., 72
rmst_tbl <- predict_flexsurvcure_metric(
  fit_list  = fits,
  times_vec = seq(0, 72, 12),
  metric    = "rmst"
)

# Compute selected quantiles
# Only quantiles below the uncured proportion are defined in cure models
quartile_tbl <- predict_flexsurvcure_metric(
  fit_list  = fits,
  metric    = "quantile",
  probs_vec = seq(0.1, 0.3, 0.1)
)

# Extract coefficients in multiple useful formats
coef_out <- extract_flexsurvcure_coefs_wide(fits)

coef_wide      <- coef_out$coef_wide
param_map_wide <- coef_out$param_map_wide
coef_long      <- coef_out$coef_long

# Print selected outputs
coef_wide
param_map_wide
coef_long
rmst_tbl

# Copy wide coefficient table to clipboard
clipr::write_clip(coef_wide)

# Optional:
# clipr::write_clip(param_map_wide)
# clipr::write_clip(quartile_tbl)
# clipr::write_clip(surv_tbl)
# clipr::write_clip(haz_tbl)
# clipr::write_clip(cumhaz_tbl)
# clipr::write_clip(rmst_tbl)