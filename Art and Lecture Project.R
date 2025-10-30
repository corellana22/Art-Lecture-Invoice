# ============================
# Invoice Analytics Loader (R)
# Robust multi-sheet loader + FY extraction + date coalescing
# ============================

library(readxl)
library(dplyr)
library(purrr)
library(stringr)
library(lubridate)
library(ggplot2)

# --- SET YOUR PATH ---
path <- "C:/Users/core3/Downloads/Data Science Projects/A&L Data Science Proj/A&L Outstanding Spreadsheet Data/Outstanding Spreadsheets Art and Lecture.xlsx"

# --- Helper: extract fiscal year from sheet names like "Outstanding Spreadsheet(2024-25"
extract_fy <- function(s) {
  m <- str_match(s, "(\\d{4})\\s*[-–]?\\s*(\\d{2}|\\d{4})")
  if (any(is.na(m))) return(NA_character_)
  y1 <- as.integer(m[,2]); y2raw <- m[,3]
  y2 <- if (nchar(y2raw) == 2) y1 + 1 else as.integer(y2raw)
  sprintf("%d-%d", y1, y2)
}

# --- Helper: guess which row contains real headers (looks for header-like keywords)
guess_header_row <- function(sheet, path, max_check = 6) {
  peek <- read_excel(path, sheet = sheet, col_names = FALSE, n_max = max_check)
  header_keywords <- c("vendor","invoice","paid","date","po","receipt","ap inbox","gateway","gus","bfs")
  scores <- sapply(1:nrow(peek), function(r) {
    vals <- tolower(trimws(as.character(unlist(peek[r, ]))))
    vals <- vals[vals != "" & !is.na(vals)]
    sum(sapply(header_keywords, function(k) any(grepl(k, vals, fixed = TRUE))))
  })
  nonempty <- sapply(1:nrow(peek), function(r) {
    sum(!is.na(peek[r, ]) & trimws(as.character(unlist(peek[r, ]))) != "")
  })
  order(scores, nonempty, decreasing = TRUE)[1]
}

# --- Helper: clean raw header names safely
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

# --- Smarter date converter: handles Date, POSIXct, numeric, and digit-only serials
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
    parsed <- suppressWarnings(parse_date_time(x_chr[need_parse],
                                               orders = c("ymd","mdy","dmy","Ymd HMS","mdy HMS","dmy HMS")))
    out[need_parse] <- as.Date(parsed)
  }
  out
}

# --- Sheets to load (only those with a recognizable FY)
all_sheets <- excel_sheets(path)
fy_labels  <- vapply(all_sheets, extract_fy, character(1))
sheets     <- all_sheets[!is.na(fy_labels)]
if (length(sheets) == 0) stop("No fiscal-year sheets recognized. Found: ", paste(all_sheets, collapse=", "))

# --- MAIN LOAD: read each FY sheet, detect header row, clean/rename columns
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
    
    # quick base coercions (we'll coalesce and re-coerce after stacking)
    if ("paid_amount" %in% names(df)) {
      x <- gsub("[^0-9.\\-]", "", as.character(df$paid_amount))
      df$paid_amount <- suppressWarnings(as.numeric(x))
    }
    df
  }
) %>% distinct()

# ============================
# Coalesce date columns across all likely sources, then compute KPIs
# ============================

# Candidate columns present in your workbook (expand if you find more)
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

as_date_safe <- function(v) if (is.null(v)) as.Date(NA) else to_date_smart(v)

coalesce_dates <- function(df, candidates) {
  cands <- candidates[candidates %in% names(df)]
  if (length(cands) == 0) return(as.Date(NA)[rep(1, nrow(df))])
  mats <- lapply(cands, function(c) as_date_safe(df[[c]]))
  out <- mats[[1]]
  if (length(mats) > 1) {
    for (k in 2:length(mats)) out <- ifelse(is.na(out), mats[[k]], out)
    out <- as.Date(out, origin = "1970-01-01")
  }
  out
}

full <- full %>%
  mutate(
    date_received_ap_inbox = coalesce_dates(., inbox_candidates),
    date_receipt_gw        = coalesce_dates(., gateway_candidates),
    gus_entry_date         = coalesce_dates(., gus_candidates),
    date_paid              = coalesce_dates(., paid_candidates),
    # recompute KPIs
    days_received_to_gw    = if (all(c("date_receipt_gw","date_received_ap_inbox") %in% names(.)))
      as.numeric(difftime(date_receipt_gw, date_received_ap_inbox, units = "days")) else NA_real_,
    days_gw_to_gus         = if (all(c("gus_entry_date","date_receipt_gw") %in% names(.)))
      as.numeric(difftime(gus_entry_date, date_receipt_gw, units = "days")) else NA_real_,
    days_received_to_paid  = if (all(c("date_paid","date_received_ap_inbox") %in% names(.)))
      as.numeric(difftime(date_paid, date_received_ap_inbox, units = "days")) else NA_real_
  )

# Convert any remaining serial/text dates to real Date
full <- full %>%
  mutate(
    date_sent_bfs = to_date_smart(date_sent_bfs)
  )

# sanity check
full %>% select(date_sent_bfs) %>% slice_head(n = 10)


# ============================
# Diagnostics
# ============================
cat("\nSheets loaded:\n"); print(sheets)
cat("\nFiscal years found:\n"); print(unique(full$fiscal_year))
cat("\nColumns:\n"); print(names(full))

cat("\nNon-missing counts for key dates:\n")
print(colSums(!is.na(dplyr::select(full, date_received_ap_inbox, date_receipt_gw, gus_entry_date, date_paid))))

cat("\nKPI summaries:\n")
print(summary(dplyr::select(full, days_received_to_gw, days_gw_to_gus, days_received_to_paid)))

# ============================
# Quick Business Plots
# ============================

# Top 10 vendors by spend
top_vendors <- full %>%
  group_by(vendor_name) %>%
  summarise(total_spend = sum(paid_amount, na.rm = TRUE), .groups = "drop") %>%
  arrange(desc(total_spend)) %>%
  slice_head(n = 10)

library(scales)

ggplot(top_vendors, aes(x = total_spend, y = reorder(vendor_name, total_spend))) +
  geom_col(fill = "#2E86AB") +
  scale_x_continuous(labels = dollar_format()) +
  labs(
    title = "Top 10 Vendors by Total Spend",
    subtitle = "Across All Fiscal Years",
    x = "Total Spend",
    y = NULL
  ) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(face="bold"),
    axis.text.y = element_text(size=12)
  )


# Lifecycle delays (AP Inbox -> Paid) by fiscal year
ggplot(full,aes(x = fiscal_year, y = days_received_to_paid, fill = fiscal_year)) +
  geom_boxplot(outliers = TRUE, outlier.color = "red") +
  scale_y_continuous(labels = comma_format()) +
  labs(
    title = "Invoice Processing Time",
    subtitle = "AP Inbox → Paid",
    x = "Fiscal Year",
    y = "Days"
  ) +
  theme_minimal(base_size=14) +
  theme(legend.position = "none")


# Total spend by fiscal year
yearly_spend <- full %>%
  group_by(fiscal_year) %>%
  summarise(total_spend = sum(paid_amount, na.rm = TRUE), .groups = "drop")

ggplot(yearly_spend, aes(x = fiscal_year, y = total_spend)) +
  geom_col(fill = "blue") +
  labs(title = "Total Spend by Fiscal Year", x = "Fiscal Year", y = "Total Spend ($)") +
  theme_minimal()

# ============================
# Advanced Analytics Practice
# ============================

library(dplyr)
library(stringr)
library(lubridate)
library(forcats)   # install.packages("forcats")
library(ggplot2)

# 0) Build modeling dataset
#    Keep rows with valid target and sensible ranges
model_df <- full %>%
  filter(!is.na(days_received_to_paid),
         days_received_to_paid >= 0,
         days_received_to_paid <= 365*2) %>%   # trim extreme outliers (optional)
  mutate(
    delay_flag = ifelse(days_received_to_paid > 30, 1L, 0L),
    # basic features
    log_amount   = log1p(paid_amount),                  # stabilizes skew
    has_po       = ifelse(is.na(po_number) | po_number=="", 0L, 1L),
    received_month = ifelse(!is.na(date_received_ap_inbox),
                            month(date_received_ap_inbox), NA_integer_),
    paid_month     = ifelse(!is.na(date_paid),
                            month(date_paid), NA_integer_)
  ) %>%
  select(delay_flag, log_amount, has_po, received_month, paid_month,
         vendor_name, fiscal_year)

# 1) Collapse rare vendors to "Other" (prevents unseen level issues & reduces dimensionality)
#    Keep top N by frequency; tune N (e.g., 30) based on your data size.
N_TOP <- 30
top_vendors <- model_df %>%
  count(vendor_name, sort = TRUE) %>%
  slice_head(n = N_TOP) %>%
  pull(vendor_name)

model_df <- model_df %>%
  mutate(
    vendor_collapsed = ifelse(vendor_name %in% top_vendors, vendor_name, "Other")
  ) %>%
  select(-vendor_name)

# 2) Set factor levels BEFORE splitting so train/test share the same space
#    Also handle fiscal_year as factor and months as factors (or leave numeric)
model_df <- model_df %>%
  mutate(
    vendor_collapsed = factor(vendor_collapsed, levels = c(top_vendors, "Other")),
    fiscal_year = factor(fiscal_year),
    received_month = factor(received_month, levels = 1:12),
    paid_month     = factor(paid_month,     levels = 1:12)
  )

# 3) Simple train/test split
set.seed(123)
idx <- sample.int(nrow(model_df), size = floor(0.7 * nrow(model_df)))
train <- model_df[idx, ]
test  <- model_df[-idx, ]

# 4) Fit logistic regression
logit_formula <- delay_flag ~ log_amount + has_po + received_month + paid_month + fiscal_year + vendor_collapsed
logit_model <- glm(logit_formula, data = train, family = binomial())

# 5) Predict on test (now vendor levels align, no error)
pred_prob <- predict(logit_model, newdata = test, type = "response")
pred_cls  <- ifelse(pred_prob > 0.5, 1L, 0L)

# 6) Quick evaluation
conf_mat <- table(Predicted = pred_cls, Actual = test$delay_flag)
conf_mat

accuracy <- sum(diag(conf_mat)) / sum(conf_mat)
accuracy

library(forecast)
library(ggplot2)

# Monthly aggregation
monthly_trend <- full %>%
  filter(!is.na(date_paid)) %>%
  mutate(month = floor_date(date_paid, "month")) %>%
  group_by(month) %>%
  summarise(total_paid = sum(paid_amount, na.rm=TRUE))

# Plot
ggplot(monthly_trend, aes(month, total_paid)) +
  geom_line(color="steelblue") +
  scale_y_continuous(labels=scales::dollar_format()) +
  labs(title="Monthly Total Paid", x="Month", y="Total Paid") +
  theme_minimal()

# Time series object
ts_data <- ts(monthly_trend$total_paid, frequency = 12,
              start = c(year(min(monthly_trend$month)),
                        month(min(monthly_trend$month))))

# Decompose trend/seasonality
decomp <- stl(ts_data, s.window="periodic")
autoplot(decomp)

# ============================
# Vendor Clustering: labeled + faceted
# ============================

# Packages (install once if needed)
# install.packages(c("dplyr","ggplot2","scales","cluster","ggrepel"))

library(dplyr)
library(ggplot2)
library(scales)
library(cluster)
library(ggrepel)

# 1) Build vendor metrics (original units)
vendor_summary <- full %>%
  group_by(vendor_name) %>%
  summarise(
    avg_spend = mean(paid_amount, na.rm = TRUE),
    avg_delay = mean(days_received_to_paid, na.rm = TRUE),
    invoices  = n(),
    .groups = "drop"
  ) %>%
  mutate(
    # mean of all-NA becomes NaN; convert to NA explicitly
    avg_delay = ifelse(is.nan(avg_delay), NA_real_, avg_delay),
    avg_spend = ifelse(is.nan(avg_spend), NA_real_, avg_spend)
  ) %>%
  # keep vendors with enough data to be meaningful
  filter(invoices >= 4)

# 2) Keep rows with at least some usable values, then impute NAs for clustering
vs <- vendor_summary %>%
  filter(is.finite(avg_spend) | is.finite(avg_delay))

# Median imputation (robust + simple)
med_spend <- median(vs$avg_spend, na.rm = TRUE)
med_delay <- median(vs$avg_delay, na.rm = TRUE)
vs_imp <- vs %>%
  mutate(
    avg_spend = ifelse(!is.finite(avg_spend), med_spend, avg_spend),
    avg_delay = ifelse(!is.finite(avg_delay), med_delay, avg_delay)
  )

# Optional: cap extreme outliers to stabilize clustering visuals
cap <- function(x, p = c(0.01, 0.99)) {
  qs <- quantile(x, probs = p, na.rm = TRUE)
  pmin(pmax(x, qs[1]), qs[2])
}
vs_imp <- vs_imp %>%
  mutate(
    avg_spend = cap(avg_spend),
    avg_delay = cap(avg_delay)
  )

# 3) Scale for k-means (only numeric features)
clust_mat <- scale(vs_imp[, c("avg_spend", "avg_delay")])

# 4) Choose k (simple elbow heuristic), then cluster
set.seed(123)
max_k <- max(2, min(6, nrow(vs_imp) - 1))  # ensure valid bounds
wss <- sapply(1:max_k, function(k) kmeans(clust_mat, centers = k, nstart = 25)$tot.withinss)
drops <- c(NA, diff(wss))
k1_wss <- wss[1]
candidate_k <- which((abs(drops) / k1_wss) < 0.10)[1]  # first small marginal gain
best_k <- ifelse(is.na(candidate_k) || candidate_k < 2, 3, candidate_k)
best_k <- min(best_k, max_k)

set.seed(123)
km <- kmeans(clust_mat, centers = best_k, nstart = 50)
vs_imp$cluster <- factor(km$cluster)

# 5) Cluster profiles (in original units) + human-readable labels
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
    spend_tier = ntile(avg_spend, 3),   # 1=low, 3=high
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

# Attach readable names to each vendor row
vs_labeled <- vs_imp %>%
  left_join(cluster_profiles %>% select(cluster, cluster_name), by = "cluster")

# 6) Centroids (original units) for labeling
centroids <- vs_labeled %>%
  group_by(cluster, cluster_name) %>%
  summarise(
    cx = mean(avg_spend, na.rm = TRUE),
    cy = mean(avg_delay, na.rm = TRUE),
    .groups = "drop"
  )

# 7A) FACET BY CLUSTER NAME (tells the story of segments)
p_clusters <- ggplot(vs_labeled, aes(x = avg_spend, y = avg_delay, color = cluster, size = invoices)) +
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
    subtitle = paste0("k = ", best_k, "  |  Vendors with ≥4 invoices; medians used to impute missing values"),
    x = "Average Spend per Invoice",
    y = "Average Days to Payment",
    color = "Cluster",
    size  = "Invoices"
  ) +
  facet_wrap(~ cluster_name, ncol = 2) +
  theme_minimal(base_size = 13) +
  theme(legend.position = "bottom")

print(p_clusters)

# ============================
# (Optional) FACET BY FISCAL YEAR (no clustering; quick comparison by FY)
# ============================

vendor_summary_fy <- full %>%
  group_by(fiscal_year, vendor_name) %>%
  summarise(
    avg_spend = mean(paid_amount, na.rm = TRUE),
    avg_delay = mean(days_received_to_paid, na.rm = TRUE),
    invoices  = n(),
    .groups = "drop"
  ) %>%
  filter(invoices >= 4)

p_fy <- ggplot(vendor_summary_fy, aes(x = avg_spend, y = avg_delay, size = invoices)) +
  geom_point(alpha = 0.6, color = "#2E86AB") +
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

print(p_fy)

summary(days_received_to_paid)
table(is.na(full$days_received_to_paid))
table(full$fiscal_year, useNA="ifany")

# Basic summary
summary(full$days_received_to_paid)

# Quantiles to choose sensible y-limits
quantile(full$days_received_to_paid, probs = c(0, .01, .25, .5, .75, .99, 1), na.rm = TRUE)

# Check for impossible values (should be none after your cleaning)
sum(full$days_received_to_paid < 0, na.rm = TRUE)

# Summary by fiscal year
full %>%
  group_by(fiscal_year) %>%
  summarise(
    n_obs   = n(),
    n_days  = sum(!is.na(days_received_to_paid)),
    median_days = median(days_received_to_paid, na.rm = TRUE),
    p75_days    = quantile(days_received_to_paid, .75, na.rm = TRUE),
    p90_days    = quantile(days_received_to_paid, .90, na.rm = TRUE)
  )

plot_data <- full %>% filter(!is.na(days_received_to_paid))
lims <- quantile(plot_data$days_received_to_paid, c(.01, .99), na.rm = TRUE)

ggplot(plot_data, aes(fiscal_year, days_received_to_paid, fill = fiscal_year)) +
  geom_boxplot(outlier.shape = 16, outlier.size = 1.3, outlier.alpha = 0.3) +
  coord_cartesian(ylim = lims) +
  labs(
    title = "Invoice Processing Time (AP Inbox → Paid)",
    subtitle = paste("n =", nrow(plot_data), "invoices with both dates"),
    x = "Fiscal Year", y = "Days"
  ) +
  theme(legend.position = "none",
        plot.title = element_text(face = "bold"))

# Histogram of days to paid (trim tails)
ggplot(plot_data, aes(days_received_to_paid)) +
  geom_histogram(bins = 40, fill = "#2E86AB") +
  coord_cartesian(xlim = lims) +
  labs(title = "Distribution of Days to Payment", x = "Days", y = "Count")

# Vendors with longest average time to payment (min 5 invoices)
vendor_delays <- plot_data %>%
  group_by(vendor_name) %>%
  summarise(avg_days = mean(days_received_to_paid), n = n(), .groups = "drop") %>%
  filter(n >= 5) %>%
  arrange(desc(avg_days)) %>%
  slice_head(n = 15)

ggplot(vendor_delays, aes(avg_days, reorder(vendor_name, avg_days))) +
  geom_col(fill = "#C0392B") +
  labs(title = "Vendors with Longest Average Days to Payment",
       subtitle = "Minimum 5 invoices",
       x = "Average Days", y = NULL)
full %>%
  summarise(
    n = n(),
    have_inbox   = sum(!is.na(date_received_ap_inbox)),
    have_gw      = sum(!is.na(date_receipt_gw)),
    have_gus     = sum(!is.na(gus_entry_date)),
    have_paid    = sum(!is.na(date_paid)),
    have_both_for_days = sum(!is.na(date_received_ap_inbox) & !is.na(date_paid))
  )




