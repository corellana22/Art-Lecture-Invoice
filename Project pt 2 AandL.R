# ============================
# Invoice Analytics Loader (R) — Refactor
# Robust multi-sheet loader + FY extraction + date coalescing + KPIs + visuals
# ============================

# ---- Packages ----
suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(purrr)
  library(stringr)
  library(lubridate)
  library(ggplot2)
  library(scales)
  library(forcats)
  library(cluster)
  library(ggrepel)
})

# ---- Config ----
cfg <- list(
  path = "C:/Users/core3/Downloads/Data Science Projects/A&L Data Science Proj/A&L Outstanding Spreadsheet Data/Outstanding Spreadsheets Art and Lecture.xlsx",
  min_invoices_for_vendor_stats = 4L,
  max_days_to_keep = 365L * 2L,   # trim extreme outliers in modeling
  delay_flag_threshold = 30L,     # >30 days = delayed
  seed = 123
)

# ---- Utils ----
extract_fy <- function(s) {
  # Accepts a single sheet name
  m <- str_match(s, "(\\d{4})\\s*[-–]?\\s*(\\d{2}|\\d{4})")
  if (any(is.na(m))) return(NA_character_)
  y1 <- as.integer(m[, 2]); y2raw <- m[, 3]
  y2 <- if (nchar(y2raw) == 2) y1 + 1L else as.integer(y2raw)
  sprintf("%d-%d", y1, y2)
}

guess_header_row <- function(sheet, path, max_check = 6) {
  peek <- read_excel(path, sheet = sheet, col_names = FALSE, n_max = max_check)
  header_keywords <- c("vendor","invoice","paid","date","po","receipt","ap inbox","gateway","gus","bfs")
  scores <- sapply(seq_len(nrow(peek)), function(r) {
    vals <- tolower(trimws(as.character(unlist(peek[r, ]))))
    vals <- vals[vals != "" & !is.na(vals)]
    sum(vapply(header_keywords, function(k) any(grepl(k, vals, fixed = TRUE)), logical(1)))
  })
  nonempty <- sapply(seq_len(nrow(peek)), function(r) {
    sum(!is.na(peek[r, ]) & trimws(as.character(unlist(peek[r, ]))) != "")
  })
  # Pick the most "header-like" row; break ties by density
  res <- order(scores, nonempty, decreasing = TRUE)[1]
  if (length(res) == 0 || is.na(res)) res <- 1
  res
}

clean_names_simple <- function(v) {
  v <- ifelse(is.na(v), "", v)
  v <- gsub("\u00A0", " ", v)        # NBSP
  v <- trimws(v)
  v <- gsub("\\s+", " ", v)
  v <- tolower(v)
  v <- gsub("[^a-z0-9]+", "_", v)
  v <- gsub("^_+|_+$", "", v)
  v[v == ""] <- paste0("empty_", seq_along(v))[v == ""]
  v
}

to_date_smart <- function(x) {
  if (inherits(x, "Date")) return(as.Date(x))
  if (inherits(x, "POSIXct") || inherits(x, "POSIXt")) return(as.Date(x))
  if (is.numeric(x)) return(as.Date(x, origin = "1899-12-30"))  # Excel serial
  
  x_chr <- trimws(as.character(x))
  x_chr[x_chr == ""] <- NA_character_
  is_digit <- !is.na(x_chr) & str_detect(x_chr, "^[0-9]+$")
  serial_num <- suppressWarnings(as.numeric(x_chr))
  looks_serial <- is_digit & !is.na(serial_num) & serial_num >= 20000 & serial_num <= 90000
  
  out <- rep(as.Date(NA), length(x_chr))
  out[looks_serial] <- as.Date(serial_num[looks_serial], origin = "1899-12-30")
  
  need_parse <- is.na(out) & !is.na(x_chr)
  if (any(need_parse)) {
    parsed <- suppressWarnings(parse_date_time(
      x_chr[need_parse],
      orders = c("ymd","mdy","dmy","Ymd HMS","mdy HMS","dmy HMS")
    ))
    out[need_parse] <- as.Date(parsed)
  }
  out
}

as_date_safe <- function(v) {
  if (is.null(v)) return(as.Date(NA))
  to_date_smart(v)
}

coalesce_dates <- function(df, candidates) {
  cands <- candidates[candidates %in% names(df)]
  if (length(cands) == 0) return(rep(as.Date(NA), nrow(df)))
  mats <- lapply(cands, function(c) as_date_safe(df[[c]]))
  out <- mats[[1]]
  if (length(mats) > 1) {
    for (k in 2:length(mats)) {
      out <- ifelse(is.na(out), mats[[k]], out)  # becomes numeric; fix next
    }
    out <- as.Date(out, origin = "1970-01-01")
  }
  out
}

numify_amount <- function(x) {
  x <- gsub("[^0-9.\\-]", "", as.character(x))
  suppressWarnings(as.numeric(x))
}

cap <- function(x, p = c(0.01, 0.99)) {
  qs <- quantile(x, probs = p, na.rm = TRUE)
  pmin(pmax(x, qs[1]), qs[2])
}

# ---- Load workbook ----
load_outstanding <- function(path) {
  if (!file.exists(path)) stop("File not found: ", path)
  all_sheets <- excel_sheets(path)
  fy_labels  <- vapply(all_sheets, extract_fy, character(1))
  sheets     <- all_sheets[!is.na(fy_labels)]
  if (length(sheets) == 0) {
    stop("No fiscal-year sheets recognized. Found: ", paste(all_sheets, collapse = ", "))
  }
  
  full <- map_dfr(
    sheets,
    ~ {
      hdr_row <- guess_header_row(.x, path)
      df <- read_excel(path, sheet = .x, skip = hdr_row - 1, .name_repair = "minimal")
      names(df) <- clean_names_simple(names(df))
      df <- df %>% select(-matches("^empty_\\d+$"), -matches("^\\.+\\d+$"))
      
      df <- df %>%
        rename(
          vendor_name              = any_of(c("vendor_name","vendor")),
          date_received_ap_inbox   = any_of(c("date_received_to_ap_inbox","date_received_ap_inbox","ap_inbox","in_box","date_received")),
          invoice_number           = any_of(c("invoice_number","invoice_no","invoice_#","invoice")),
          paid_amount              = any_of(c("paid_amount","amount","invoice_amount","total")),
          gus_entry_date           = any_of(c("gus_entry_date","gus_date","date_gus")),
          date_receipt_gw          = any_of(c("date_receipt_in_gw","date_receipt_gw","gateway_receipt_date","for_gateway","gateway_date")),
          po_number                = any_of(c("po_number","po","po_")),
          receipt_number           = any_of(c("receipt_number","receipt","receipt_")),
          date_sent_bfs            = any_of(c("date_sent_to_bfs_invoices_only","date_sent_to_bfs","date_sent_bfs")),
          date_paid                = any_of(c("date_paid_gateway_dw_gl","date_paid","paid_date","paid_finance_analyst","invoices_only_marked_paid_in_gateway"))
        ) %>%
        mutate(fiscal_year = extract_fy(.x))
      
      if ("paid_amount" %in% names(df)) df$paid_amount <- numify_amount(df$paid_amount)
      df
    }
  ) %>% distinct()
  
  attr(full, "sheets_loaded") <- sheets
  full
}

# ---- Build canonical dates + KPIs ----
canonicalize_dates <- function(full) {
  inbox_candidates <- c(
    "date_received_ap_inbox","date_received_to_ap_inbox",
    "in_box","in_box_finance_analyst_student_finance_assistant","date_received"
  )
  gateway_candidates <- c(
    "date_receipt_gw","date_receipt_in_gw",
    "for_gateway","for_gateway_finance_analyst","gateway_date","gateway_receipt_date"
  )
  gus_candidates  <- c("gus_entry_date","date_gus","gus_date")
  paid_candidates <- c(
    "date_paid","date_paid_gateway_dw_gl","paid_date",
    "paid_finance_analyst","invoices_only_marked_paid_in_gateway"
  )
  
  full %>%
    mutate(
      date_received_ap_inbox = coalesce_dates(., inbox_candidates),
      date_receipt_gw        = coalesce_dates(., gateway_candidates),
      gus_entry_date         = coalesce_dates(., gus_candidates),
      date_paid              = coalesce_dates(., paid_candidates),
      date_sent_bfs          = to_date_smart(date_sent_bfs),
      days_received_to_gw    = if (all(c("date_receipt_gw","date_received_ap_inbox") %in% names(.)))
        as.numeric(difftime(date_receipt_gw, date_received_ap_inbox, units = "days")) else NA_real_,
      days_gw_to_gus         = if (all(c("gus_entry_date","date_receipt_gw") %in% names(.)))
        as.numeric(difftime(gus_entry_date, date_receipt_gw, units = "days")) else NA_real_,
      days_received_to_paid  = if (all(c("date_paid","date_received_ap_inbox") %in% names(.)))
        as.numeric(difftime(date_paid, date_received_ap_inbox, units = "days")) else NA_real_
    )
}

# ---- Diagnostics ----
print_diagnostics <- function(full, sheets) {
  cat("\nSheets loaded:\n"); print(sheets)
  cat("\nFiscal years found:\n"); print(unique(full$fiscal_year))
  cat("\nColumns:\n"); print(names(full))
  cat("\nNon-missing counts for key dates:\n")
  print(colSums(!is.na(dplyr::select(full, date_received_ap_inbox, date_receipt_gw, gus_entry_date, date_paid))))
  cat("\nKPI summaries:\n")
  print(summary(dplyr::select(full, days_received_to_gw, days_gw_to_gus, days_received_to_paid)))
}

# ---- Plots (call on demand) ----
plot_top_spenders <- function(full, n = 10) {
  top_spenders <- full %>%
    group_by(vendor_name) %>%
    summarise(total_spend = sum(paid_amount, na.rm = TRUE), .groups = "drop") %>%
    arrange(desc(total_spend)) %>%
    slice_head(n = n)
  
  ggplot(top_spenders, aes(x = total_spend, y = reorder(vendor_name, total_spend))) +
    geom_col() +
    scale_x_continuous(labels = dollar_format()) +
    labs(
      title = paste0("Top ", n, " Vendors by Total Spend"),
      subtitle = "Across All Fiscal Years",
      x = "Total Spend",
      y = NULL
    ) +
    theme_minimal(base_size = 14) +
    theme(plot.title = element_text(face="bold"), axis.text.y = element_text(size=12))
}

plot_processing_time_by_fy <- function(full) {
  ggplot(full, aes(x = fiscal_year, y = days_received_to_paid, fill = fiscal_year)) +
    geom_boxplot(outliers = TRUE) +
    scale_y_continuous(labels = comma_format()) +
    labs(
      title = "Invoice Processing Time",
      subtitle = "AP Inbox → Paid",
      x = "Fiscal Year",
      y = "Days"
    ) +
    theme_minimal(base_size=14) +
    theme(legend.position = "none")
}

plot_yearly_spend <- function(full) {
  yearly_spend <- full %>%
    group_by(fiscal_year) %>%
    summarise(total_spend = sum(paid_amount, na.rm = TRUE), .groups = "drop")
  
  ggplot(yearly_spend, aes(x = fiscal_year, y = total_spend)) +
    geom_col() +
    scale_y_continuous(labels = dollar_format()) +
    labs(title = "Total Spend by Fiscal Year", x = "Fiscal Year", y = "Total Spend") +
    theme_minimal(base_size = 14)
}

# ---- Modeling dataset + logistic baseline ----
build_model_frame <- function(full, cfg) {
  full %>%
    filter(!is.na(days_received_to_paid),
           days_received_to_paid >= 0,
           days_received_to_paid <= cfg$max_days_to_keep) %>%
    mutate(
      delay_flag = ifelse(days_received_to_paid > cfg$delay_flag_threshold, 1L, 0L),
      log_amount   = log1p(paid_amount),
      has_po       = ifelse(is.na(po_number) | po_number == "", 0L, 1L),
      received_month = ifelse(!is.na(date_received_ap_inbox), month(date_received_ap_inbox), NA_integer_),
      paid_month     = ifelse(!is.na(date_paid),               month(date_paid),               NA_integer_)
    ) %>%
    select(delay_flag, log_amount, has_po, received_month, paid_month,
           vendor_name, fiscal_year)
}

fit_logistic_baseline <- function(model_df, cfg, N_TOP = 30) {
  # Collapse rare vendors
  top_vendor_levels <- model_df %>%
    count(vendor_name, sort = TRUE) %>%
    slice_head(n = N_TOP) %>%
    pull(vendor_name)
  
  model_df <- model_df %>%
    mutate(
      vendor_collapsed = ifelse(vendor_name %in% top_vendor_levels, vendor_name, "Other"),
      vendor_collapsed = factor(vendor_collapsed, levels = c(top_vendor_levels, "Other")),
      fiscal_year = factor(fiscal_year),
      received_month = factor(received_month, levels = as.character(1:12)),
      paid_month     = factor(paid_month,     levels = as.character(1:12))
    ) %>%
    select(-vendor_name)
  
  set.seed(cfg$seed)
  idx <- sample.int(nrow(model_df), size = floor(0.7 * nrow(model_df)))
  train <- model_df[idx, ]
  test  <- model_df[-idx, ]
  
  logit_formula <- delay_flag ~ log_amount + has_po + received_month + paid_month + fiscal_year + vendor_collapsed
  logit_model <- glm(logit_formula, data = train, family = binomial())
  
  pred_prob <- predict(logit_model, newdata = test, type = "response")
  pred_cls  <- ifelse(pred_prob > 0.5, 1L, 0L)
  
  conf_mat <- table(Predicted = pred_cls, Actual = test$delay_flag)
  accuracy <- sum(diag(conf_mat)) / sum(conf_mat)
  
  list(model = logit_model, conf_mat = conf_mat, accuracy = accuracy,
       train = train, test = test)
}

# ---- Vendor clustering + FY facets ----
vendor_clustering_plot <- function(full, cfg) {
  vendor_summary <- full %>%
    group_by(vendor_name) %>%
    summarise(
      avg_spend = mean(paid_amount, na.rm = TRUE),
      avg_delay = mean(days_received_to_paid, na.rm = TRUE),
      invoices  = n(),
      .groups = "drop"
    ) %>%
    mutate(
      avg_delay = ifelse(is.nan(avg_delay), NA_real_, avg_delay),
      avg_spend = ifelse(is.nan(avg_spend), NA_real_, avg_spend)
    ) %>%
    filter(invoices >= cfg$min_invoices_for_vendor_stats)
  
  vs <- vendor_summary %>%
    filter(is.finite(avg_spend) | is.finite(avg_delay))
  if (nrow(vs) < 3) stop("Not enough vendors for clustering after filtering.")
  
  med_spend <- median(vs$avg_spend, na.rm = TRUE)
  med_delay <- median(vs$avg_delay, na.rm = TRUE)
  vs_imp <- vs %>%
    mutate(
      avg_spend = ifelse(!is.finite(avg_spend), med_spend, avg_spend),
      avg_delay = ifelse(!is.finite(avg_delay), med_delay, avg_delay),
      avg_spend = cap(avg_spend),
      avg_delay = cap(avg_delay)
    )
  
  clust_mat <- scale(vs_imp[, c("avg_spend", "avg_delay")])
  set.seed(cfg$seed)
  max_k <- max(2, min(6, nrow(vs_imp) - 1))
  wss <- sapply(1:max_k, function(k) kmeans(clust_mat, centers = k, nstart = 25)$tot.withinss)
  drops <- c(NA, diff(wss))
  k1_wss <- wss[1]
  candidate_k <- which((abs(drops) / k1_wss) < 0.10)[1]
  best_k <- ifelse(is.na(candidate_k) || candidate_k < 2, 3, candidate_k)
  best_k <- min(best_k, max_k)
  
  set.seed(cfg$seed)
  km <- kmeans(clust_mat, centers = best_k, nstart = 50)
  vs_imp$cluster <- factor(km$cluster)
  
  cluster_profiles <- vs_imp %>%
    group_by(cluster) %>%
    summarise(
      vendors   = n(),
      avg_spend = mean(avg_spend, na.rm = TRUE),
      avg_delay = mean(avg_delay, na.rm = TRUE),
      invoices  = sum(invoices, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(
      spend_tier = ntile(avg_spend, 3),
      delay_tier = ntile(avg_delay, 3),
      cluster_name = dplyr::case_when(
        spend_tier == 3 & delay_tier == 3 ~ "High spend / Long delay",
        spend_tier == 3 & delay_tier == 2 ~ "High spend / Medium delay",
        spend_tier == 3 & delay_tier == 1 ~ "High spend / Fast pay",
        spend_tier == 2 & delay_tier == 3 ~ "Mid spend / Long delay",
        spend_tier == 2 & delay_tier == 2 ~ "Mid spend / Medium delay",
        spend_tier == 2 & delay_tier == 1 ~ "Mid spend / Fast pay",
        spend_tier == 1 & delay_tier == 3 ~ "Low spend / Long delay",
        spend_tier == 1 & delay_tier == 2 ~ "Low spend / Medium delay",
        TRUE                               ~ "Low spend / Fast pay"
      )
    )
  
  vs_labeled <- vs_imp %>%
    left_join(cluster_profiles %>% select(cluster, cluster_name), by = "cluster")
  
  centroids <- vs_labeled %>%
    group_by(cluster, cluster_name) %>%
    summarise(
      cx = mean(avg_spend, na.rm = TRUE),
      cy = mean(avg_delay, na.rm = TRUE),
      .groups = "drop"
    )
  
  ggplot(vs_labeled, aes(x = avg_spend, y = avg_delay, color = cluster, size = invoices)) +
    geom_point(alpha = 0.75) +
    geom_point(data = centroids, aes(cx, cy), color = "black", fill = "white",
               shape = 21, size = 4, inherit.aes = FALSE) +
    ggrepel::geom_label_repel(
      data = centroids,
      aes(cx, cy, label = cluster_name),
      color = "black", fill = "white", size = 3.5, inherit.aes = FALSE, max.overlaps = 30
    ) +
    scale_x_continuous(labels = dollar_format()) +
    labs(
      title = "Vendor Clusters by Spend & Payment Delay",
      subtitle = paste0("k = ", best_k, "  |  Vendors with ≥", cfg$min_invoices_for_vendor_stats, " invoices; medians imputed"),
      x = "Average Spend per Invoice",
      y = "Average Days to Payment",
      color = "Cluster",
      size  = "Invoices"
    ) +
    facet_wrap(~ cluster_name, ncol = 2) +
    theme_minimal(base_size = 13) +
    theme(legend.position = "bottom")
}

facet_fy_scatter <- function(full, cfg) {
  vendor_summary_fy <- full %>%
    group_by(fiscal_year, vendor_name) %>%
    summarise(
      avg_spend = mean(paid_amount, na.rm = TRUE),
      avg_delay = mean(days_received_to_paid, na.rm = TRUE),
      invoices  = n(),
      .groups = "drop"
    ) %>%
    filter(invoices >= cfg$min_invoices_for_vendor_stats)
  
  ggplot(vendor_summary_fy, aes(x = avg_spend, y = avg_delay, size = invoices)) +
    geom_point(alpha = 0.6) +
    scale_x_continuous(labels = dollar_format()) +
    labs(
      title = "Vendor Spend vs. Delay by Fiscal Year",
      x = "Average Spend per Invoice",
      y = "Average Days to Payment",
      size = "Invoices"
    ) +
    facet_wrap(~ fiscal_year) +
    theme_minimal(base_size = 13) +
    theme(legend.position = "bottom")
}

# ---- Main ----
main <- function(cfg) {
  full_raw <- load_outstanding(cfg$path)
  full <- canonicalize_dates(full_raw)
  
  # Diagnostics
  print_diagnostics(full, attr(full_raw, "sheets_loaded"))
  
  # Quick summaries
  message("\nBasic summary of days_received_to_paid:")
  print(summary(full$days_received_to_paid))
  message("\n1%/25%/50%/75%/99% quantiles (trim tails):")
  print(quantile(full$days_received_to_paid, probs = c(0, .01, .25, .5, .75, .99, 1), na.rm = TRUE))
  message("\nNegative days (should be 0 after cleaning):")
  print(sum(full$days_received_to_paid < 0, na.rm = TRUE))
  
  message("\nFY-level stats:")
  print(
    full %>%
      group_by(fiscal_year) %>%
      summarise(
        n_obs   = n(),
        n_days  = sum(!is.na(days_received_to_paid)),
        median_days = median(days_received_to_paid, na.rm = TRUE),
        p75_days    = quantile(days_received_to_paid, .75, na.rm = TRUE),
        p90_days    = quantile(days_received_to_paid, .90, na.rm = TRUE),
        .groups = "drop"
      )
  )
  
  # Plots (uncomment to render)
print(plot_top_spenders(full, n = 10))
print(plot_processing_time_by_fy(full))
print(plot_yearly_spend(full))
print(vendor_clustering_plot(full, cfg))
print(facet_fy_scatter(full, cfg))
  
  # Modeling baseline
  model_df <- build_model_frame(full, cfg)
  if (nrow(model_df) >= 30) {
    mdl <- fit_logistic_baseline(model_df, cfg)
    message("\nConfusion matrix:\n"); print(mdl$conf_mat)
    message(sprintf("\nAccuracy: %.3f", mdl$accuracy))
  } else {
    message("\n[Modeling skipped] Not enough rows with valid target.")
  }
  
  invisible(full)
}

# Run
full_result <- main(cfg)
