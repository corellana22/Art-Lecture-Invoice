# ============================
# Shiny Dashboard (uses existing `full` with days_received_to_paid)
# ============================
library(shiny)
library(shinydashboard)   # <--- this provides valueBox, valueBoxOutput, etc.
library(bslib)
library(tidyverse)
library(lubridate)
library(scales)
library(ggrepel)

# ---- sanity: require `full` in memory ----
stopifnot(exists("full"), is.data.frame(full))

# order fiscal years nicely
full <- full %>%
  mutate(fiscal_year = factor(fiscal_year, levels = sort(unique(fiscal_year))))

# helper: inline "clean" view of days without mutating data
clean_days <- function(x) {
  ifelse(is.na(x) | x < 0 | x > 365*2, NA_real_, x)
}
on_time_flag <- function(x, thr = 30) ifelse(!is.na(x) & x <= thr, 1L, 0L)

# ---- UI ----
ui <- page_sidebar(
  theme = bs_theme(bootswatch = "flatly", base_font = font_google("Inter")),
  title = "Invoice Lifecycle Analytics",
  sidebar = sidebar(
    h4("Filters"),
    selectizeInput("fy", "Fiscal year",
                   choices = levels(full$fiscal_year),
                   selected = levels(full$fiscal_year),
                   multiple = TRUE),
    selectizeInput("vendor", "Vendor (optional)",
                   choices = sort(unique(full$vendor_name)),
                   selected = NULL, multiple = TRUE,
                   options = list(placeholder = 'All vendors')),
    dateRangeInput("paid_range", "Paid date range",
                   start = suppressWarnings(min(full$date_paid, na.rm = TRUE)),
                   end   = suppressWarnings(max(full$date_paid, na.rm = TRUE))),
    numericInput("min_invoices", "Min invoices per vendor", value = 3, min = 1, step = 1),
    sliderInput("sla_thr", "On-time threshold (days)", min = 7, max = 60, value = 30),
    actionButton("reset", "Reset filters", class = "btn-primary")
  ),
  layout_columns(
    col_widths = c(3,3,3,3),
    valueBoxOutput("v_total_spend"),
    valueBoxOutput("v_invoices"),
    valueBoxOutput("v_median_days"),
    valueBoxOutput("v_on_time")
  ),
  navset_pill(
    nav_panel("Trends",
              plotOutput("p_monthly", height = 350),
              br(),
              plotOutput("p_spend_fy", height = 350)
    ),
    nav_panel("Lifecycle",
              plotOutput("p_box", height = 380),
              br(),
              plotOutput("p_hist", height = 300)
    ),
    nav_panel("Vendors",
              plotOutput("p_vendor_spend", height = 380),
              br(),
              plotOutput("p_vendor_delay", height = 380),
              br(),
              tableOutput("tbl_vendor_sla")
    ),
    nav_panel("Clusters (Spend vs Delay)",
              helpText("Simple k-means using avg spend & delay per vendor; computed on-the-fly."),
              plotOutput("p_clusters", height = 420)
    )
  ),
  footer = div(class="text-muted small",
               "Data: Outstanding Spreadsheets (A&L). Built with Shiny.")
)

# ---- SERVER ----
server <- function(input, output, session) {
  
  observeEvent(input$reset, {
    updateSelectizeInput(session, "fy", selected = levels(full$fiscal_year))
    updateSelectizeInput(session, "vendor", selected = character(0))
    updateDateRangeInput(session, "paid_range",
                         start = suppressWarnings(min(full$date_paid, na.rm=TRUE)),
                         end   = suppressWarnings(max(full$date_paid, na.rm=TRUE)))
    updateNumericInput(session, "min_invoices", value = 3)
    updateSliderInput(session, "sla_thr", value = 30)
  })
  
  # Filtered data
  d <- reactive({
    x <- full %>% filter(fiscal_year %in% input$fy)
    if (!is.null(input$paid_range[1]) && !is.null(input$paid_range[2])) {
      x <- x %>%
        filter(is.na(date_paid) | (date_paid >= input$paid_range[1] & date_paid <= input$paid_range[2]))
    }
    if (length(input$vendor)) x <- x %>% filter(vendor_name %in% input$vendor)
    x
  })
  
  # KPIs
  output$v_total_spend <- renderValueBox({
    total_spend <- d() |>
      summarise(sum = sum(paid_amount, na.rm = TRUE)) |>
      pull(sum)
    shinydashboard::valueBox(
      value    = scales::dollar(total_spend, accuracy = 1),
      subtitle = "Total Spend",
      color    = "blue"
    )
  })
  
  output$v_invoices <- renderValueBox({
    shinydashboard::valueBox(
      value    = scales::comma(nrow(d())),
      subtitle = "Invoices",
      color    = "blue"
    )
  })
  
  output$v_median_days <- renderValueBox({
    med <- d() |>
      summarise(m = median(ifelse(is.na(days_received_to_paid) |
                                    days_received_to_paid < 0 |
                                    days_received_to_paid > 730, NA, days_received_to_paid),
                           na.rm = TRUE)) |>
      pull(m)
    shinydashboard::valueBox(
      value    = ifelse(is.finite(med), round(med, 1), "—"),
      subtitle = "Median Days to Paid",
      color    = "blue"
    )
  })
  
  output$v_on_time <- renderValueBox({
    thr <- input$sla_thr
    pct <- d() |>
      summarise(p = mean(ifelse(!is.na(days_received_to_paid) &
                                  days_received_to_paid <= thr, 1L, 0L),
                         na.rm = TRUE) * 100) |>
      pull(p)
    shinydashboard::valueBox(
      value    = ifelse(is.finite(pct), paste0(round(pct, 1), "%"), "—"),
      subtitle = paste0("On-time \u2264 ", thr, " days"),
      color    = "blue"
    )
  })
  
  # ---- Trends ----
  output$p_monthly <- renderPlot({
    d() %>%
      filter(!is.na(date_paid)) %>%
      mutate(month = floor_date(date_paid, "month")) %>%
      summarise(total_paid = sum(paid_amount, na.rm = TRUE), .by = month) %>%
      complete(month = seq(min(month), max(month), by = "month"), fill = list(total_paid = 0)) %>%
      ggplot(aes(month, total_paid)) +
      geom_line(linewidth = 1) +
      scale_y_continuous(labels = label_dollar()) +
      labs(title = "Monthly Total Paid", x = "Month", y = "Total")
  })
  
  output$p_spend_fy <- renderPlot({
    d() %>%
      summarise(total = sum(paid_amount, na.rm = TRUE), .by = fiscal_year) %>%
      ggplot(aes(fiscal_year, total)) +
      geom_col(fill = "#2E86AB") +
      scale_y_continuous(labels = label_dollar()) +
      labs(title = "Total Spend by Fiscal Year", x = NULL, y = "Total Spend")
  })
  
  # ---- Lifecycle ----
  output$p_box <- renderPlot({
    dd <- d() %>% mutate(days = clean_days(days_received_to_paid)) %>% filter(!is.na(days))
    if (!nrow(dd)) return(NULL)
    lims <- quantile(dd$days, c(.01, .99), na.rm = TRUE)
    ggplot(dd, aes(fiscal_year, days, fill = fiscal_year)) +
      geom_boxplot(outlier.shape = 16, outlier.size = 1.3, outlier.alpha = 0.3) +
      coord_cartesian(ylim = lims) +
      labs(title = "Processing Time (AP Inbox → Paid)", x = "Fiscal Year", y = "Days") +
      theme_minimal(base_size = 14) + theme(legend.position = "none")
  })
  
  output$p_hist <- renderPlot({
    dd <- d() %>% mutate(days = clean_days(days_received_to_paid)) %>% filter(!is.na(days))
    if (!nrow(dd)) return(NULL)
    lims <- quantile(dd$days, c(.01, .99), na.rm = TRUE)
    ggplot(dd, aes(days)) +
      geom_histogram(bins = 40, fill = "#2E86AB") +
      coord_cartesian(xlim = lims) +
      labs(title = "Distribution of Days to Payment", x = "Days", y = "Count")
  })
  
  # ---- Vendors ----
  output$p_vendor_spend <- renderPlot({
    dd <- d() %>%
      summarise(total = sum(paid_amount, na.rm = TRUE), .by = vendor_name) %>%
      filter(total > 0) %>%
      slice_max(total, n = 12)
    if (!nrow(dd)) return(NULL)
    ggplot(dd, aes(total, reorder(vendor_name, total))) +
      geom_col(fill="#2E86AB") +
      scale_x_continuous(labels = label_dollar()) +
      labs(title = "Top Vendors by Spend", x = "Total Spend", y = NULL)
  })
  
  output$p_vendor_delay <- renderPlot({
    dd <- d() %>%
      mutate(days = clean_days(days_received_to_paid)) %>%
      filter(!is.na(days)) %>%
      summarise(avg_days = mean(days), n = dplyr::n(), .by = vendor_name) %>%
      filter(n >= input$min_invoices) %>%
      slice_max(avg_days, n = 12)
    if (!nrow(dd)) return(NULL)
    ggplot(dd, aes(avg_days, reorder(vendor_name, avg_days))) +
      geom_col(fill="#C0392B") +
      labs(title = "Slowest Vendors (Avg Days to Paid)",
           subtitle = paste("Min", input$min_invoices, "invoices"),
           x = "Average Days", y = NULL)
  })
  
  output$tbl_vendor_sla <- renderTable({
    d() %>%
      mutate(days = clean_days(days_received_to_paid)) %>%
      filter(!is.na(days)) %>%
      mutate(on_time = on_time_flag(days, input$sla_thr)) %>%
      summarise(
        invoices = dplyr::n(),
        total_spend = sum(paid_amount, na.rm = TRUE),
        median_days = median(days),
        pct_on_time = mean(on_time)*100,
        .by = c(fiscal_year)
      ) %>%
      arrange(fiscal_year) %>%
      mutate(total_spend = dollar(total_spend),
             pct_on_time = paste0(round(pct_on_time, 1), "%"))
  })
  
  # ---- Clusters (quick) ----
  output$p_clusters <- renderPlot({
    dd <- d() %>%
      mutate(days = clean_days(days_received_to_paid)) %>%
      summarise(
        avg_spend = mean(paid_amount, na.rm = TRUE),
        avg_delay = mean(days, na.rm = TRUE),
        invoices  = dplyr::n(), .by = vendor_name
      ) %>%
      filter(invoices >= input$min_invoices,
             is.finite(avg_spend), is.finite(avg_delay))
    if (nrow(dd) < 3) return(NULL)
    
    # scale + simple k=3 (or fewer if data is tiny)
    X <- scale(dd[, c("avg_spend","avg_delay")])
    set.seed(123)
    km <- kmeans(X, centers = min(3, nrow(dd)-1), nstart = 25)
    dd$cluster <- factor(km$cluster)
    
    cents <- dd %>% group_by(cluster) %>% summarise(cx=mean(avg_spend), cy=mean(avg_delay), .groups="drop")
    
    ggplot(dd, aes(avg_spend, avg_delay, color = cluster, size = invoices)) +
      geom_point(alpha = 0.75) +
      geom_point(data = cents, aes(cx, cy), color="black", fill="white",
                 shape=21, size=4, inherit.aes = FALSE) +
      ggrepel::geom_label_repel(data = cents, aes(cx, cy, label = paste("Cluster", cluster)),
                                color="black", fill="white", size=3.5, inherit.aes = FALSE) +
      scale_x_continuous(labels = label_dollar()) +
      labs(title = "Vendor Clusters (Avg Spend vs Avg Delay)",
           x = "Avg Spend per Invoice", y = "Avg Days to Paid",
           color = "Cluster", size = "Invoices") +
      theme(legend.position = "bottom")
  })
}

shinyApp(ui, server)
