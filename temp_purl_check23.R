## ----setup, include=FALSE-----------------------------------------------------
knitr::opts_chunk$set(
  echo = TRUE,
  dpi = 96,
  fig.retina = 1
)


## ----warning=FALSE, message=FALSE, echo=FALSE---------------------------------
setwd("E:/ARC/Projects/A1C/data/all/arc_a1c")
# remove(a1c_final)
# load("a1c_final.Rdata")
# load("us_states.Rdata")
# load("geo_usa_select.Rdata")
# load("nhanesdat.Rdata")
# library(gtsummary)
library(huxtable)
library(sumExtras)
library(dplyr)
library(cowplot)
library(ggplot2)
library(htmltools)
library(htmlwidgets)
library(leaflet)
library(msm)
library(sf)
library(spdep)
library(tigris)
library(downloadthis)
library(readr)
library(readxl)
library(lubridate)
library(ggpubr)
library(RColorBrewer)
library(knitr)
library(DT)
library(gt)
# library(plotly)
library(leaflegend)
library(ggfx)
library(ggiraph)
# library(mapview)
library(tigris)
library(flextable)
library(scales)
# library(crosstalk)
# Install and load the package
# install.packages("showtext")
library(showtext)
library(huxtable)
library(leaflet.extras2)
library(openxlsx2)
library(flexlsx)
library(data.table)
library(gtsummary)
library(gt)
library(stringr)
customeTheme = theme(
  title = element_text(size = 11),
  axis.title.x = element_text(size = 12),
  axis.text.x = element_text(size = 12),
  axis.title.y = element_text(size = 12))

load("a1c_all.Rdata")
load("a1c_all_a1cdemo.Rdata")

medical_history_cols <- c(
  "cancer", "cardiorespiratory", "coagulopathy", "teratogen",
  "anticoagulants", "deferral_meds", "hypertension", "unwell",
  "splenectomy_tcpenia", "read_edu_mat"
)

a1c_all_a1cdemo <- a1c_all_a1cdemo %>%
  dplyr::mutate(
    `Donation Date` = as.Date(.data[["Donation Date"]]),
    SBP = as.numeric(stringr::str_match(.data[["BP"]], "^(\\d{2,3})/(\\d{2,3})$")[, 2]),
    DBP = as.numeric(stringr::str_match(.data[["BP"]], "^(\\d{2,3})/(\\d{2,3})$")[, 3]),
    BP_category = dplyr::case_when(
      is.na(SBP) & is.na(DBP) ~ NA_character_,
      (!is.na(SBP) & SBP > 180) | (!is.na(DBP) & DBP > 120) ~ "Hypertensive crisis",
      (!is.na(SBP) & SBP >= 140) | (!is.na(DBP) & DBP >= 90) ~ "Hypertension stage 2",
      (!is.na(SBP) & SBP >= 130) | (!is.na(DBP) & DBP >= 80) ~ "Hypertension stage 1",
      !is.na(SBP) & !is.na(DBP) & SBP >= 120 & SBP < 130 & DBP < 80 ~ "Elevated",
      !is.na(SBP) & !is.na(DBP) & SBP < 120 & DBP < 80 ~ "Normal",
      TRUE ~ NA_character_
    ),
    BP_category = factor(
      BP_category,
      levels = c(
        "Normal",
        "Elevated",
        "Hypertension stage 1",
        "Hypertension stage 2",
        "Hypertensive crisis"
      )
    )
  ) %>%
  dplyr::mutate(
    dplyr::across(
      dplyr::all_of(medical_history_cols),
      ~ ifelse(is.na(.x), 0, .x)
    ),
    `ARC A1C Results` = as.numeric(.data[["ARC A1C Results"]]),
    glycemic_category = dplyr::case_when(
      is.na(.data[["ARC A1C Results"]]) ~ NA_character_,
      .data[["ARC A1C Results"]] < 5.7 ~ "Normoglycemia",
      .data[["ARC A1C Results"]] >= 5.7 & .data[["ARC A1C Results"]] < 6.5 ~ "Prediabetes",
      .data[["ARC A1C Results"]] >= 6.5 ~ "Diabetes"
    ),
    glycemic_category = factor(
      glycemic_category,
      levels = c("Normoglycemia", "Prediabetes", "Diabetes")
    ),
    period = dplyr::case_when(
      .data[["Donation Date"]] >= as.Date("2025-03-01") & .data[["Donation Date"]] <= as.Date("2025-04-30") ~ "Mar 2025",
      .data[["Donation Date"]] >= as.Date("2025-07-01") & .data[["Donation Date"]] <= as.Date("2025-09-30") ~ "Aug 2025",
      .data[["Donation Date"]] >= as.Date("2025-11-01") & .data[["Donation Date"]] <= as.Date("2025-12-31") ~ "Nov 2025",
      .data[["Donation Date"]] >= as.Date("2026-03-01") & .data[["Donation Date"]] <= as.Date("2026-04-30") ~ "Mar 2026",
      TRUE ~ NA_character_
    )
  ) %>% rename(a1c = `ARC A1C Results`) |>
  mutate(raceeth = case_when(race %in% c("Caucasian") ~ "White",
                                                      race %in% c("African American") ~ "Black",
                                                      race %in% c("Hispanic") ~ "Hispanic Origin",
                                                      race %in% c("Asian","Mix","Native American", "Other", "Prefer not to answer","") ~ "Asian/AIAN/other",
                                                      TRUE~""
)
)

a1c_all_a1cdemo_dedup <- a1c_all_a1cdemo %>%
  dplyr::arrange(.data[["id"]], dplyr::desc(.data[["Donation Date"]])) %>%
  dplyr::distinct(.data[["id"]], .keep_all = TRUE) 

a1c_repeat_daily <- a1c_all_a1cdemo %>%
  dplyr::filter(!is.na(a1c), !is.na(id), !is.na(.data[["Donation Date"]])) %>%
  dplyr::group_by(id, `Donation Date`) %>%
  dplyr::summarise(
    a1c = max(a1c, na.rm = TRUE),
    gender = dplyr::first(gender[!is.na(gender) & gender != ""], default = NA_character_),
    raceeth = dplyr::first(raceeth[!is.na(raceeth) & raceeth != ""], default = NA_character_),
    zip_raw = dplyr::first(
      as.character(zip)[!is.na(zip) & trimws(as.character(zip)) != ""],
      default = NA_character_
    ),
    age = if (all(is.na(age))) NA_real_ else max(age, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  dplyr::mutate(
    zip5 = stringr::str_pad(
      stringr::str_extract(zip_raw, "\\d{1,5}"),
      width = 5,
      side = "left",
      pad = "0"
    ),
    glycemic_category = dplyr::case_when(
      a1c < 5.7 ~ "Normoglycemia",
      a1c >= 5.7 & a1c < 6.5 ~ "Prediabetes",
      a1c >= 6.5 ~ "Diabetes"
    ),
    glycemic_category = factor(
      glycemic_category,
      levels = c("Normoglycemia", "Prediabetes", "Diabetes")
    )
  ) %>%
  dplyr::select(-zip_raw)

a1c_repeat_subset <- a1c_repeat_daily %>%
  dplyr::arrange(id, .data[["Donation Date"]]) %>%
  dplyr::add_count(id, name = "n_obs") %>%
  dplyr::filter(n_obs > 1) %>%
  dplyr::group_by(id) %>%
  dplyr::mutate(
    visit_num = dplyr::row_number(),
    interval_days = as.numeric(.data[["Donation Date"]] - dplyr::lag(.data[["Donation Date"]])),
    interval_months = interval_days / 30.4375,
    cum_years = as.numeric(.data[["Donation Date"]] - dplyr::first(.data[["Donation Date"]])) / 365.25,
    status = dplyr::case_when(
      glycemic_category == "Normoglycemia" ~ 1L,
      glycemic_category == "Prediabetes" ~ 2L,
      glycemic_category == "Diabetes" ~ 3L
    ),
    gender_male = dplyr::case_when(
      gender == "M" ~ 1,
      gender == "F" ~ 0,
      TRUE ~ NA_real_
    ),
    black_indicator = dplyr::case_when(
      raceeth == "Black" ~ 1,
      is.na(raceeth) ~ NA_real_,
      TRUE ~ 0
    )
  ) %>%
  dplyr::ungroup()

repeat_interval_summary <- a1c_repeat_subset %>%
  dplyr::filter(!is.na(interval_days), interval_days > 0) %>%
  dplyr::summarise(
    transitions = dplyr::n(),
    mean_days = mean(interval_days, na.rm = TRUE),
    median_days = median(interval_days, na.rm = TRUE),
    q1_days = quantile(interval_days, 0.25, na.rm = TRUE),
    q3_days = quantile(interval_days, 0.75, na.rm = TRUE)
  )

repeat_donor_summary <- tibble::tibble(
  metric = c(
    "Repeat donors",
    "Repeat donor-date observations",
    "Positive follow-up intervals",
    "Median interval between A1c results (days)",
    "Mean interval between A1c results (days)"
  ),
  value = c(
    dplyr::n_distinct(a1c_repeat_subset$id),
    nrow(a1c_repeat_subset),
    repeat_interval_summary$transitions,
    round(repeat_interval_summary$median_days, 1),
    round(repeat_interval_summary$mean_days, 1)
  )
)

repeat_transition_counts <- msm::statetable.msm(
  status,
  id,
  data = a1c_repeat_subset %>% dplyr::select(id, status, cum_years)
)

set.seed(42)
repeat_model_ids <- sample(
  unique(a1c_repeat_subset$id),
  size = min(5000, dplyr::n_distinct(a1c_repeat_subset$id))
)

a1c_repeat_model_data <- a1c_repeat_subset %>%
  dplyr::filter(id %in% repeat_model_ids)

repeat_q_init <- rbind(
  c(0, 0.10, 0.02),
  c(0.06, 0, 0.12),
  c(0.03, 0.04, 0)
)

a1c_repeat_msm <- msm::msm(
  status ~ cum_years,
  subject = id,
  data = a1c_repeat_model_data,
  qmatrix = repeat_q_init,
  deathexact = FALSE,
  exacttimes = TRUE,
  control = list(trace = 0, REPORT = 1, maxit = 500)
)

repeat_qmatrix <- msm::qmatrix.msm(a1c_repeat_msm)
repeat_pmatrix_1y <- msm::pmatrix.msm(a1c_repeat_msm, t = 1, ci = "none")
repeat_state_labels <- c("Normoglycemia", "Prediabetes", "Diabetes")

repeat_tidy_qmatrix <- function(qmatrix_obj) {
  q_estimate <- if (is.list(qmatrix_obj) && !is.null(qmatrix_obj$estimate)) {
    qmatrix_obj$estimate
  } else {
    qmatrix_obj
  }

  as.data.frame(as.table(as.matrix(q_estimate)), responseName = "intensity") %>%
    dplyr::rename(from_state_raw = Var1, to_state_raw = Var2) %>%
    dplyr::mutate(
      from_id = as.integer(from_state_raw),
      to_id = as.integer(to_state_raw)
    ) %>%
    dplyr::filter(from_id != to_id) %>%
    dplyr::mutate(
      from_state = dplyr::recode(
        as.character(from_state_raw),
        `1` = "Normoglycemia",
        `State 1` = "Normoglycemia",
        `2` = "Prediabetes",
        `State 2` = "Prediabetes",
        `3` = "Diabetes",
        `State 3` = "Diabetes",
        .default = as.character(from_state_raw)
      ),
      to_state = dplyr::recode(
        as.character(to_state_raw),
        `1` = "Normoglycemia",
        `State 1` = "Normoglycemia",
        `2` = "Prediabetes",
        `State 2` = "Prediabetes",
        `3` = "Diabetes",
        `State 3` = "Diabetes",
        .default = as.character(to_state_raw)
      ),
      transition = paste(from_state, "->", to_state)
    ) %>%
    dplyr::select(from_id, to_id, transition, intensity)
}

repeat_race_msm <- tryCatch(
  msm::msm(
    status ~ cum_years,
    subject = id,
    data = a1c_repeat_model_data %>%
      dplyr::filter(!is.na(black_indicator)),
    qmatrix = repeat_q_init,
    deathexact = FALSE,
    exacttimes = TRUE,
    covariates = ~ black_indicator,
    constraint = list(black_indicator = c(1, 2, 3, 1, 2, 3)),
    control = list(trace = 0, REPORT = 1, maxit = 500)
  ),
  error = function(e) e
)

repeat_gender_msm <- tryCatch(
  msm::msm(
    status ~ cum_years,
    subject = id,
    data = a1c_repeat_model_data %>%
      dplyr::filter(!is.na(gender_male)),
    qmatrix = repeat_q_init,
    deathexact = FALSE,
    exacttimes = TRUE,
    covariates = ~ gender_male,
    constraint = list(gender_male = c(1, 2, 3, 1, 2, 3)),
    control = list(trace = 0, REPORT = 1, maxit = 500)
  ),
  error = function(e) e
)

a1c_all <- a1c_all %>%
  dplyr::mutate(
    `Donation Date` = as.Date(.data[["Donation Date"]]),
    `ARC A1C Results` = as.numeric(.data[["ARC A1C Results"]]),
    glycemic_category = dplyr::case_when(
      is.na(.data[["ARC A1C Results"]]) ~ NA_character_,
      .data[["ARC A1C Results"]] < 5.7 ~ "Normoglycemia",
      .data[["ARC A1C Results"]] >= 5.7 & .data[["ARC A1C Results"]] < 6.5 ~ "Prediabetes",
      .data[["ARC A1C Results"]] >= 6.5 ~ "Diabetes"
    ),
    glycemic_category = factor(
      glycemic_category,
      levels = c("Normoglycemia", "Prediabetes", "Diabetes")
    ),
    period = dplyr::case_when(
      .data[["Donation Date"]] >= as.Date("2025-03-01") & .data[["Donation Date"]] <= as.Date("2025-04-30") ~ "Mar 2025",
      .data[["Donation Date"]] >= as.Date("2025-07-01") & .data[["Donation Date"]] <= as.Date("2025-09-30") ~ "Aug 2025",
      .data[["Donation Date"]] >= as.Date("2025-11-01") & .data[["Donation Date"]] <= as.Date("2025-12-31") ~ "Nov 2025",
      .data[["Donation Date"]] >= as.Date("2026-03-01") & .data[["Donation Date"]] <= as.Date("2026-04-30") ~ "Mar 2026",
      TRUE ~ NA_character_
    )
  )

period_donor_counts <- a1c_all %>%
  dplyr::filter(!is.na(period)) %>%
  dplyr::distinct(period, id) %>%
  dplyr::count(period, name = "distinct_donors") %>%
  dplyr::mutate(
    period = factor(period, levels = c("Mar 2025", "Aug 2025", "Nov 2025", "Mar 2026"))
  ) %>%
  dplyr::arrange(period)

period_plot_counts <- period_donor_counts

glycemic_period_counts <- a1c_all %>%
  dplyr::filter(!is.na(period), !is.na(glycemic_category)) %>%
  dplyr::arrange(
    .data[["period"]],
    .data[["id"]],
    .data[["Donation Date"]],
    .data[["Donation Id"]]
  ) %>%
  dplyr::distinct(period, id, .keep_all = TRUE) %>%
  dplyr::count(period, glycemic_category, name = "distinct_donors") %>%
  dplyr::group_by(period) %>%
  dplyr::mutate(
    period_total = sum(distinct_donors),
    pct = distinct_donors / period_total
  ) %>%
  dplyr::ungroup() %>%
  dplyr::mutate(
    period = factor(period, levels = c("Mar 2025", "Aug 2025", "Nov 2025", "Mar 2026"))
  )

repeat_donors_2025_2026 <- a1c_all %>%
  dplyr::filter(period %in% c("Mar 2025", "Aug 2025", "Nov 2025")) %>%
  dplyr::distinct(id) %>%
  dplyr::inner_join(
    a1c_all %>%
      dplyr::filter(period == "Mar 2026") %>%
      dplyr::distinct(id),
    by = "id"
  ) %>%
  nrow()

ids_2025 <- a1c_all %>%
  dplyr::filter(period %in% c("Mar 2025", "Aug 2025", "Nov 2025")) %>%
  dplyr::distinct(id)

ids_2026 <- a1c_all %>%
  dplyr::filter(period == "Mar 2026") %>%
  dplyr::distinct(id)

mar_2026_repeat_share <- repeat_donors_2025_2026 /
  (period_plot_counts %>%
     dplyr::filter(period == "Mar 2026") %>%
     dplyr::pull(distinct_donors))

retention_share_2025_to_2026 <- repeat_donors_2025_2026 / nrow(ids_2025)

retention_flow <- ids_2025 %>%
  dplyr::mutate(
    `2025 cohort` = "Total donors 2025",
    `Mar 2026 status` = dplyr::if_else(
      id %in% ids_2026[["id"]],
      "Rescreened in Mar 2026",
      "Not yet rescreened"
    )
  ) %>%
  dplyr::count(`2025 cohort`, `Mar 2026 status`, name = "donors") %>%
  dplyr::mutate(
    `Mar 2026 status` = factor(
      `Mar 2026 status`,
      levels = c("Rescreened in Mar 2026", "Not yet rescreened")
    )
  ) %>%
  dplyr::arrange(`Mar 2026 status`)

period_donor_counts <- dplyr::bind_rows(
  period_donor_counts %>%
    dplyr::mutate(period = as.character(period)),
  tibble::tibble(
    period = "Repeat donors (2025 to Mar 2026)",
    distinct_donors = repeat_donors_2025_2026
  )
)


## ----donor-period-counts, echo=FALSE, warning=FALSE, message=FALSE, echo=FALSE----
knitr::kable(period_donor_counts, col.names = c("Screening months", "Distinct donors"))


## ----donor-period-visual, echo=FALSE, fig.width=11, fig.height=5, warning=FALSE, message=FALSE, echo=FALSE----
period_bar <- ggplot(
  period_plot_counts,
  aes(x = period, y = distinct_donors, fill = period)
) +
  geom_col(width = 0.65, show.legend = FALSE) +
  geom_text(
    aes(label = scales::comma(distinct_donors)),
    vjust = -0.35,
    size = 4
  ) +
  scale_fill_manual(
    values = c(
      "Mar 2025" = "#C9A66B",
      "Aug 2025" = "#B78352",
      "Nov 2025" = "#8D5E3C",
      "Mar 2026" = "#1F6C73"
    )
  ) +
  scale_y_continuous(
    labels = scales::comma,
    expand = expansion(mult = c(0, 0.14))
  ) +
  labs(
    title = "Distinct donors by screening months",
    x = NULL,
    y = "Distinct donors"
  ) +
  theme_minimal(base_size = 12) +
  theme(
    panel.grid.minor = element_blank(),
    panel.grid.major.x = element_blank(),
    axis.text.x = element_text(size = 11),
    plot.title.position = "plot"
  )

repeat_callout <- ggplot() +
  annotate("rect", xmin = 0, xmax = 1, ymin = 0, ymax = 1, fill = "#F4F1EA", color = NA) +
  annotate(
    "text",
    x = 0.5,
    y = 0.68,
    label = paste0("Repeat donors\n", scales::comma(repeat_donors_2025_2026)),
    fontface = "bold",
    size = 6,
    color = "#17324D"
  ) +
  annotate(
    "text",
    x = 0.5,
    y = 0.38,
    label = paste0(
      scales::percent(mar_2026_repeat_share, accuracy = 0.1),
      " of Mar 2026 donors\nwere previously screened in 2025"
    ),
    size = 4.2,
    color = "#4B5563"
  ) +
  coord_cartesian(xlim = c(0, 1), ylim = c(0, 1), expand = FALSE) +
  theme_void()

cowplot::plot_grid(period_bar, repeat_callout, rel_widths = c(3.2, 1.4))


## ----warning=FALSE, message=FALSE, echo=FALSE---------------------------------
make_missing_last <- function(x, levels = NULL) {
  x_chr <- as.character(x)
  x_chr[is.na(x_chr) | trimws(x_chr) == ""] <- "Missing"

  if (is.null(levels)) {
    levels <- c(unique(x_chr[x_chr != "Missing"]), "Missing")
  } else {
    levels <- c(levels, "Missing")
  }

  factor(x_chr, levels = unique(levels))
}

t0_data <- a1c_all_a1cdemo_dedup %>%
  dplyr::mutate(
    gender = make_missing_last(gender, c("F", "M")),
    raceeth = make_missing_last(
      raceeth,
      c("White", "Black", "Hispanic Origin", "Asian/AIAN/other")
    ),
    age_group = make_missing_last(age_group, c("-29", "30-54", "55-")),
    ftd = make_missing_last(ftd, c("N", "Y")),
    glycemic_category = make_missing_last(
      glycemic_category,
      c("Normoglycemia", "Prediabetes", "Diabetes")
    ),
    BP_category = make_missing_last(
      BP_category,
      c(
        "Normal",
        "Elevated",
        "Hypertension stage 1",
        "Hypertension stage 2",
        "Hypertensive crisis"
      )
    )
  )

t0 <- t0_data %>% tbl_summary(
                                         include = c(gender, raceeth, age, age_group, ftd, age, glycemic_category, Hb, SBP, DBP, BP_category ,cancer, hypertension, cardiorespiratory, coagulopathy, teratogen, anticoagulants, deferral_meds, splenectomy_tcpenia
                                                     ),
                                         label = list(age_group = "age group",
                                                      ftd = "first-time donor",
                                                      # hispanic_flag = "hispanic flag",
                                                      # BMI_cat = "BMI category",
                                                      median_household_income = "median household income",
                                                      raceeth = "race/ethnicity",
                                                      glycemic_category = "glycemic status",
                                                      BP_category = "BP range",
                                                      Hb = "hemoglobin%",
                                                      cardiorespiratory = "heart/lung condition",
                                                      teratogen = "on teratogens",
                                                      anticoagulants = "on anticoagulants",
                                                      deferral_meds = "on deferral medications",
                                                      splenectomy_tcpenia = "h/o splenectomy/thrombocytopenia",
                                                      hypertension = "h/o hypertension"
                                                      
                                         ), missing = "no",
                                         percent = "column" )%>% add_stat_label() %>%
  add_variable_group_header(
    header = "Medical History",
    variables = c(
      "cancer", "hypertension", "cardiorespiratory", "coagulopathy",
      "teratogen", "anticoagulants", "deferral_meds", "splenectomy_tcpenia"
    )
  ) %>%
  modify_caption("Donor Demographics")


medical_history_vars <- c(
  "cancer", "hypertension", "cardiorespiratory", "coagulopathy",
  "teratogen", "anticoagulants", "deferral_meds", "splenectomy_tcpenia"
)
gray_header_rows <- which(
  t0$table_body$row_type == "label" &
    !t0$table_body$variable %in% medical_history_vars
)
medical_history_header_rows <- which(t0$table_body$label == "Medical History")
medical_history_detail_rows <- which(t0$table_body$variable %in% medical_history_vars)

t0_gt <- t0 %>%
  as_gt() %>%
  theme_gt_compact() %>%
  gt::tab_style(
    style = list(
      gt::cell_fill(color = "#F2F2F2"),
      gt::cell_borders(
        sides = "top",
        color = "black",
        weight = gt::px(1)
      )
    ),
    locations = gt::cells_body(rows = gray_header_rows)
  ) %>%
  gt::tab_style(
    style = list(
      gt::cell_fill(color = "#F2F2F2"),
      gt::cell_borders(
        sides = "top",
        color = "black",
        weight = gt::px(1)
      )
    ),
    locations = gt::cells_body(rows = medical_history_header_rows)
  ) %>%
  gt::tab_style(
    style = gt::cell_borders(
      sides = "top",
      color = "#D9D9D9",
      weight = gt::px(1)
    ),
    locations = gt::cells_body(rows = medical_history_detail_rows)
  )


t0_gt_excel <- t0$table_body %>%
  dplyr::select(label, stat_0) %>%
  dplyr::rename(
    Characteristic = label,
    `N = 920,112` = stat_0
  )

t0_gt_excel_path <- file.path(getwd(), "t0_gt_table.xlsx")
t0_gt_wb <- openxlsx::createWorkbook()
openxlsx::addWorksheet(t0_gt_wb, "t0_gt")
openxlsx::writeData(t0_gt_wb, "t0_gt", t0_gt_excel)
openxlsx::setColWidths(t0_gt_wb, "t0_gt", cols = 1:ncol(t0_gt_excel), widths = "auto")
openxlsx::addStyle(
  t0_gt_wb, "t0_gt",
  style = openxlsx::createStyle(textDecoration = "bold"),
  rows = 1, cols = 1:ncol(t0_gt_excel), gridExpand = TRUE, stack = TRUE
)
openxlsx::addStyle(
  t0_gt_wb, "t0_gt",
  style = openxlsx::createStyle(
    fgFill = "#F2F2F2",
    border = "top",
    borderColour = "#000000",
    borderStyle = "thin"
  ),
  rows = gray_header_rows + 1,
  cols = 1:ncol(t0_gt_excel),
  gridExpand = TRUE, stack = TRUE
)
openxlsx::addStyle(
  t0_gt_wb, "t0_gt",
  style = openxlsx::createStyle(
    fgFill = "#F2F2F2",
    border = "top",
    borderColour = "#000000",
    borderStyle = "thin"
  ),
  rows = medical_history_header_rows + 1,
  cols = 1:ncol(t0_gt_excel),
  gridExpand = TRUE, stack = TRUE
)
openxlsx::addStyle(
  t0_gt_wb, "t0_gt",
  style = openxlsx::createStyle(
    border = "top",
    borderColour = "#D9D9D9",
    borderStyle = "thin"
  ),
  rows = medical_history_detail_rows + 1,
  cols = 1:ncol(t0_gt_excel),
  gridExpand = TRUE, stack = TRUE
)
openxlsx::saveWorkbook(t0_gt_wb, t0_gt_excel_path, overwrite = TRUE)

download_file( 
  path = t0_gt_excel_path,
  output_name = "2025-2026 A1c Screening Donor Summary",
  button_label = "download table (Excel)" ,
  button_type = "info",
  has_icon = TRUE,
  self_contained = FALSE
  )


t0_gt


## ----glycemic-category-by-period, echo=FALSE, fig.width=10, fig.height=5.5, warning=FALSE, message=FALSE, echo=FALSE----
ggplot(
  glycemic_period_counts,
  aes(x = period, y = pct, fill = glycemic_category)
) +
  geom_col(width = 0.7, color = "white", position = position_stack(reverse = TRUE)) +
  geom_text(
    aes(
      label = ifelse(
        pct >= 0.045 | glycemic_category == "Diabetes",
        scales::percent(pct, accuracy = 0.1),
        ""
      )
    ),
    position = position_stack(vjust = 0.5, reverse = TRUE),
    size = 3.6,
    color = "white",
    fontface = "bold"
  ) +
  scale_fill_manual(
    values = c(
      "Normoglycemia" = "#5B8E7D",
      "Prediabetes" = "#D39C3F",
      "Diabetes" = "#B44B3E"
    )
  ) +
  scale_y_continuous(
    labels = scales::percent,
    expand = expansion(mult = c(0, 0.02))
  ) +
  labs(
    title = "Glycemic category distribution by screening months",
    subtitle = "Distinct donors within each screening month group; percentages based on each group total",
    x = NULL,
    y = "Percentage of distinct donors",
    fill = "Glycemic category"
  ) +
  theme_minimal(base_size = 12) +
  theme(
    panel.grid.minor = element_blank(),
    panel.grid.major.x = element_blank(),
    axis.text.x = element_text(size = 11),
    legend.position = "top",
    plot.title.position = "plot"
  ) +
  guides(fill = guide_legend(reverse = TRUE))


## ----urban-rural-strat-prep, warning=FALSE, message=FALSE, include=FALSE------
urban_rural_strat_data <- tryCatch({
  options(tigris_use_cache = TRUE)

  format_county_geoid <- function(x) {
    x_chr <- trimws(as.character(x))
    x_chr <- sub("\\.0$", "", x_chr)
    stringr::str_pad(x_chr, width = 5, side = "left", pad = "0")
  }

  urban_rural_county_lookup <- readxl::read_excel("urban rural classification CDC.xlsx") %>%
    {
      names(.)[names(.) == "2023 Code"] <- "urban_code"
      .
    } %>%
    dplyr::transmute(
      GEOID = format_county_geoid(Location),
      urban_rural_group = dplyr::case_when(
        substr(as.character(urban_code), 1, 1) %in% c("1", "2", "3", "4") ~ "Urban",
        substr(as.character(urban_code), 1, 1) %in% c("5", "6") ~ "Rural",
        TRUE ~ "Unknown"
      )
    )

  donor_geo <- a1c_all_a1cdemo_dedup %>%
    dplyr::filter(
      !is.na(age_group),
      !is.na(glycemic_category),
      glycemic_category != "hypoglycemia"
    ) %>%
    dplyr::mutate(
      zip5 = stringr::str_pad(
        stringr::str_extract(as.character(zip), "\\d{1,5}"),
        width = 5,
        side = "left",
        pad = "0"
      )
    ) %>%
    dplyr::filter(!is.na(zip5))

  zip_prefixes <- sort(unique(substr(donor_geo$zip5, 1, 3)))

  zcta_sf <- tigris::zctas(
    cb = TRUE,
    starts_with = zip_prefixes,
    year = 2020,
    class = "sf"
  )

  counties_sf <- tigris::counties(
    cb = TRUE,
    year = 2020,
    class = "sf"
  ) %>%
    dplyr::filter(!STATEFP %in% c("60", "66", "69", "72", "78"))

  zip_id_col <- intersect(c("ZCTA5CE20", "GEOID20", "ZCTA5CE10"), names(zcta_sf))[1]

  zcta_to_county_lookup <- zcta_sf %>%
    sf::st_transform(5070) %>%
    sf::st_point_on_surface() %>%
    sf::st_transform(sf::st_crs(counties_sf)) %>%
    sf::st_join(counties_sf %>% dplyr::select(GEOID), left = FALSE) %>%
    sf::st_drop_geometry() %>%
    dplyr::transmute(
      zip5 = .data[[zip_id_col]],
      GEOID = .data[["GEOID"]]
    ) %>%
    dplyr::distinct(zip5, .keep_all = TRUE)

  donor_geo %>%
    dplyr::left_join(zcta_to_county_lookup, by = "zip5") %>%
    dplyr::left_join(urban_rural_county_lookup, by = "GEOID") %>%
    dplyr::filter(urban_rural_group %in% c("Urban", "Rural")) %>%
    dplyr::mutate(
      urban_rural_group = factor(
        urban_rural_group,
        levels = c("Urban", "Rural")
      )
    )
}, error = function(e) {
  list(error = conditionMessage(e))
})


## ----warning=FALSE, message=FALSE, echo=FALSE---------------------------------
df_a1ccount <- a1c_all_a1cdemo_dedup %>% filter(!is.na(gender)) %>%
  group_by(age_group, gender) %>% 
  summarise(donor_count=n(),mean_A1c = round(mean(a1c, na.rm= T), digits = 2))
df_a1ccatcount <- a1c_all_a1cdemo_dedup %>%  filter(!is.na(gender)) %>%
  group_by(age_group, gender, glycemic_category) %>% count()
df_count <- a1c_all_a1cdemo_dedup %>%  filter(!is.na(gender)) %>%
  group_by(age_group, gender) %>% count()

df_diab_p <- df_a1ccatcount %>% left_join(df_count, by = c("age_group", "gender"))  |>
  mutate(percentage =round(n.x/n.y, digits = 3)) |>
  rename(`group count` = n.x, `total donors` = n.y)
df_diab_perc <- df_diab_p %>% mutate(percent = paste0(round(100*percentage, digits = 1), "%")) %>% select(-percentage) 

df_diab_p <- df_diab_p %>% filter(glycemic_category != "hypoglycemia")
glycemic_tab <- xtabs(
  ~ interaction(age_group, gender) + glycemic_category,
  data = a1c_all_a1cdemo_dedup %>%
    filter(!is.na(age_group), !is.na(gender), !is.na(glycemic_category), glycemic_category != "hypoglycemia") %>%
    mutate(glycemic_category = droplevels(factor(glycemic_category)))
)
glycemic_chisq <- chisq.test(glycemic_tab)
glycemic_prop <- 100 * prop.table(glycemic_tab, 1)
r <- ggplot(df_diab_p, aes(x=age_group, y = percentage)) + 
  geom_bar(aes(color = glycemic_category, fill=glycemic_category), stat = "identity", position = position_dodge2(0.8), width = 0.7 ) +
  facet_wrap(~gender) +
  geom_text( aes(label = paste0(100*percentage, "%"), group = glycemic_category),
             position = position_dodge(0.8),
             vjust = 0, hjust = -0.1, size = 3.5) +
  coord_flip() + 
  scale_fill_brewer(palette = "YlOrRd", direction = 1) +
  scale_color_brewer(palette = "YlOrRd", direction = 1) + guides(fill = guide_legend(reverse = T)) + 
  guides(color = guide_legend(reverse = T)) +
  scale_y_continuous(labels = scales::percent, limits = c(0,1.05)) +
  labs(title="Glycemic status of donors by age group stratified by gender",
       subtitle = "among donors screened in 2025-2026",
       caption = "donors with A1c% below 4.0 are excluded",
       x = "donor age group",
       y = "percentage of donors") + 
  customeTheme +
  theme(axis.text.x = element_blank(), axis.ticks.x = element_blank(),
        axis.ticks.y = element_blank(),
        axis.text.y = element_text(angle = 0, vjust = 0.5, hjust=0.5),
        title = element_text(size = 11),
        legend.position = "right",
        rect = element_rect(fill = "transparent"),
        panel.grid.major.y = element_blank(),
        panel.grid.minor.y = element_blank(),text=element_text(size=12),
        panel.background = element_rect(fill = "transparent", colour = NA), # Panel background
        plot.background = element_rect(fill = "transparent", colour = NA),  # Plot background
        panel.grid.major = element_blank(), # Remove major grid lines
        panel.grid.minor = element_blank(), # Remove minor grid lines
  ) 
# +  guides(fill = guide_legend(reverse = T))
r


## ----glycemic-by-raceeth, warning=FALSE, message=FALSE, echo=FALSE------------
race_plot_data <- a1c_all_a1cdemo_dedup %>%
  dplyr::filter(
    !is.na(age_group),
    !is.na(glycemic_category),
    glycemic_category != "hypoglycemia",
    !is.na(raceeth),
    trimws(raceeth) != ""
  ) %>%
  dplyr::mutate(
    raceeth = factor(
      raceeth,
      levels = c("White", "Black", "Hispanic Origin", "Asian/AIAN/other")
    )
  ) %>%
  dplyr::filter(!is.na(raceeth)) %>%
  dplyr::group_by(age_group, raceeth, glycemic_category) %>%
  dplyr::summarise(group_count = dplyr::n(), .groups = "drop") %>%
  dplyr::left_join(
    a1c_all_a1cdemo_dedup %>%
      dplyr::filter(
        !is.na(age_group),
        !is.na(glycemic_category),
        glycemic_category != "hypoglycemia",
        !is.na(raceeth),
        trimws(raceeth) != ""
      ) %>%
      dplyr::mutate(
        raceeth = factor(
          raceeth,
          levels = c("White", "Black", "Hispanic Origin", "Asian/AIAN/other")
        )
      ) %>%
      dplyr::filter(!is.na(raceeth)) %>%
      dplyr::group_by(age_group, raceeth) %>%
      dplyr::summarise(total_donors = dplyr::n(), .groups = "drop"),
    by = c("age_group", "raceeth")
  ) %>%
  dplyr::mutate(percentage = group_count / total_donors)

race_glycemic_tab <- xtabs(
  ~ interaction(age_group, raceeth) + glycemic_category,
  data = a1c_all_a1cdemo_dedup %>%
    dplyr::filter(
      !is.na(age_group),
      !is.na(glycemic_category),
      glycemic_category != "hypoglycemia",
      !is.na(raceeth),
      trimws(raceeth) != ""
    ) %>%
    dplyr::mutate(
      raceeth = factor(
        raceeth,
        levels = c("White", "Black", "Hispanic Origin", "Asian/AIAN/other")
      ),
      glycemic_category = droplevels(factor(glycemic_category))
    ) %>%
    dplyr::filter(!is.na(raceeth))
)

race_glycemic_chisq <- chisq.test(race_glycemic_tab)

ggplot(race_plot_data, aes(x = age_group, y = percentage)) +
  geom_bar(
    aes(color = glycemic_category, fill = glycemic_category),
    stat = "identity",
    position = position_dodge2(0.8),
    width = 0.7
  ) +
  facet_wrap(~ raceeth, ncol = 2) +
  geom_text(
    aes(
      label = paste0(round(100 * percentage, 1), "%"),
      group = glycemic_category
    ),
    position = position_dodge(0.8),
    vjust = 0,
    hjust = -0.1,
    size = 3.3
  ) +
  coord_flip() +
  scale_fill_brewer(palette = "YlOrRd", direction = 1) +
  scale_color_brewer(palette = "YlOrRd", direction = 1) +
  guides(fill = guide_legend(reverse = TRUE)) +
  guides(color = guide_legend(reverse = TRUE)) +
  scale_y_continuous(labels = scales::percent, limits = c(0, 1.05)) +
  labs(
    title = "Glycemic status of donors by age group stratified by race/ethnicity",
    subtitle = paste(
      "among donors screened in 2025-2026;",
      "chi-square p",
      scales::pvalue(race_glycemic_chisq$p.value, accuracy = 0.001)
    ),
    caption = "donors with A1c% below 4.0 are excluded",
    x = "donor age group",
    y = "percentage of donors"
  ) +
  customeTheme +
  theme(
    axis.text.x = element_blank(),
    axis.ticks.x = element_blank(),
    axis.ticks.y = element_blank(),
    axis.text.y = element_text(angle = 0, vjust = 0.5, hjust = 0.5),
    title = element_text(size = 11),
    legend.position = "right",
    rect = element_rect(fill = "transparent"),
    panel.grid.major.y = element_blank(),
    panel.grid.minor.y = element_blank(),
    text = element_text(size = 12),
    panel.background = element_rect(fill = "transparent", colour = NA),
    plot.background = element_rect(fill = "transparent", colour = NA),
    panel.grid.major = element_blank(),
    panel.grid.minor = element_blank()
  )


## ----glycemic-by-donor-type, warning=FALSE, message=FALSE, echo=FALSE---------
donor_type_plot_data <- a1c_all_a1cdemo_dedup %>%
  dplyr::filter(
    !is.na(age_group),
    !is.na(glycemic_category),
    glycemic_category != "hypoglycemia",
    !is.na(ftd)
  ) %>%
  dplyr::mutate(
    donor_type = dplyr::case_when(
      ftd == "Y" ~ "First-time donor",
      ftd == "N" ~ "Repeat donor",
      TRUE ~ NA_character_
    ),
    donor_type = factor(
      donor_type,
      levels = c("First-time donor", "Repeat donor")
    )
  ) %>%
  dplyr::filter(!is.na(donor_type)) %>%
  dplyr::group_by(age_group, donor_type, glycemic_category) %>%
  dplyr::summarise(group_count = dplyr::n(), .groups = "drop") %>%
  dplyr::left_join(
    a1c_all_a1cdemo_dedup %>%
      dplyr::filter(
        !is.na(age_group),
        !is.na(glycemic_category),
        glycemic_category != "hypoglycemia",
        !is.na(ftd)
      ) %>%
      dplyr::mutate(
        donor_type = dplyr::case_when(
          ftd == "Y" ~ "First-time donor",
          ftd == "N" ~ "Repeat donor",
          TRUE ~ NA_character_
        ),
        donor_type = factor(
          donor_type,
          levels = c("First-time donor", "Repeat donor")
        )
      ) %>%
      dplyr::filter(!is.na(donor_type)) %>%
      dplyr::group_by(age_group, donor_type) %>%
      dplyr::summarise(total_donors = dplyr::n(), .groups = "drop"),
    by = c("age_group", "donor_type")
  ) %>%
  dplyr::mutate(percentage = group_count / total_donors)

donor_type_glycemic_tab <- xtabs(
  ~ interaction(age_group, donor_type) + glycemic_category,
  data = a1c_all_a1cdemo_dedup %>%
    dplyr::filter(
      !is.na(age_group),
      !is.na(glycemic_category),
      glycemic_category != "hypoglycemia",
      !is.na(ftd)
    ) %>%
    dplyr::mutate(
      donor_type = dplyr::case_when(
        ftd == "Y" ~ "First-time donor",
        ftd == "N" ~ "Repeat donor",
        TRUE ~ NA_character_
      ),
      donor_type = factor(
        donor_type,
        levels = c("First-time donor", "Repeat donor")
      ),
      glycemic_category = droplevels(factor(glycemic_category))
    ) %>%
    dplyr::filter(!is.na(donor_type))
)

donor_type_glycemic_chisq <- chisq.test(donor_type_glycemic_tab)

ggplot(donor_type_plot_data, aes(x = age_group, y = percentage)) +
  geom_bar(
    aes(color = glycemic_category, fill = glycemic_category),
    stat = "identity",
    position = position_dodge2(0.8),
    width = 0.7
  ) +
  facet_wrap(~ donor_type) +
  geom_text(
    aes(
      label = paste0(round(100 * percentage, 1), "%"),
      group = glycemic_category
    ),
    position = position_dodge(0.8),
    vjust = 0,
    hjust = -0.1,
    size = 3.5
  ) +
  coord_flip() +
  scale_fill_brewer(palette = "YlOrRd", direction = 1) +
  scale_color_brewer(palette = "YlOrRd", direction = 1) +
  guides(fill = guide_legend(reverse = TRUE)) +
  guides(color = guide_legend(reverse = TRUE)) +
  scale_y_continuous(labels = scales::percent, limits = c(0, 1.05)) +
  labs(
    title = "Glycemic status of donors by age group stratified by donor type",
    subtitle = paste(
      "among donors screened in 2025-2026;",
      "chi-square p",
      scales::pvalue(donor_type_glycemic_chisq$p.value, accuracy = 0.001)
    ),
    caption = "donors with A1c% below 4.0 are excluded",
    x = "donor age group",
    y = "percentage of donors"
  ) +
  customeTheme +
  theme(
    axis.text.x = element_blank(),
    axis.ticks.x = element_blank(),
    axis.ticks.y = element_blank(),
    axis.text.y = element_text(angle = 0, vjust = 0.5, hjust = 0.5),
    title = element_text(size = 11),
    legend.position = "right",
    rect = element_rect(fill = "transparent"),
    panel.grid.major.y = element_blank(),
    panel.grid.minor.y = element_blank(),
    text = element_text(size = 12),
    panel.background = element_rect(fill = "transparent", colour = NA),
    plot.background = element_rect(fill = "transparent", colour = NA),
    panel.grid.major = element_blank(),
    panel.grid.minor = element_blank()
  )


## ----glycemic-by-urban-rural, warning=FALSE, message=FALSE, echo=FALSE--------
if (inherits(urban_rural_strat_data, "list") && !is.null(urban_rural_strat_data$error)) {
  htmltools::tags$div(
    style = "padding: 12px; border: 1px solid #dddddd; background: #fafafa;",
    htmltools::tags$strong("Urban-rural stratified plot unavailable."),
    htmltools::tags$div(paste("Reason:", urban_rural_strat_data$error))
  )
} else {
  urban_rural_plot_data <- urban_rural_strat_data %>%
    dplyr::group_by(age_group, urban_rural_group, glycemic_category) %>%
    dplyr::summarise(group_count = dplyr::n(), .groups = "drop") %>%
    dplyr::left_join(
      urban_rural_strat_data %>%
        dplyr::group_by(age_group, urban_rural_group) %>%
        dplyr::summarise(total_donors = dplyr::n(), .groups = "drop"),
      by = c("age_group", "urban_rural_group")
    ) %>%
    dplyr::mutate(percentage = group_count / total_donors)

  urban_rural_glycemic_tab <- xtabs(
    ~ interaction(age_group, urban_rural_group) + glycemic_category,
    data = urban_rural_strat_data %>%
      dplyr::mutate(glycemic_category = droplevels(factor(glycemic_category)))
  )

  urban_rural_glycemic_chisq <- chisq.test(urban_rural_glycemic_tab)

  ggplot(urban_rural_plot_data, aes(x = age_group, y = percentage)) +
    geom_bar(
      aes(color = glycemic_category, fill = glycemic_category),
      stat = "identity",
      position = position_dodge2(0.8),
      width = 0.7
    ) +
    facet_wrap(~ urban_rural_group) +
    geom_text(
      aes(
        label = paste0(round(100 * percentage, 1), "%"),
        group = glycemic_category
      ),
      position = position_dodge(0.8),
      vjust = 0,
      hjust = -0.1,
      size = 3.5
    ) +
    coord_flip() +
    scale_fill_brewer(palette = "YlOrRd", direction = 1) +
    scale_color_brewer(palette = "YlOrRd", direction = 1) +
    guides(fill = guide_legend(reverse = TRUE)) +
    guides(color = guide_legend(reverse = TRUE)) +
    scale_y_continuous(labels = scales::percent, limits = c(0, 1.05)) +
    labs(
      title = "Glycemic status of donors by age group stratified by county urban-rural type",
      subtitle = paste(
        "among donors screened in 2025-2026;",
        "chi-square p",
        scales::pvalue(urban_rural_glycemic_chisq$p.value, accuracy = 0.001)
      ),
      caption = "donors with A1c% below 4.0 are excluded; county type assigned from donor ZIP linked to county urban-rural classification",
      x = "donor age group",
      y = "percentage of donors"
    ) +
    customeTheme +
    theme(
      axis.text.x = element_blank(),
      axis.ticks.x = element_blank(),
      axis.ticks.y = element_blank(),
      axis.text.y = element_text(angle = 0, vjust = 0.5, hjust = 0.5),
      title = element_text(size = 11),
      legend.position = "right",
      rect = element_rect(fill = "transparent"),
      panel.grid.major.y = element_blank(),
      panel.grid.minor.y = element_blank(),
      text = element_text(size = 12),
      panel.background = element_rect(fill = "transparent", colour = NA),
      plot.background = element_rect(fill = "transparent", colour = NA),
      panel.grid.major = element_blank(),
      panel.grid.minor = element_blank()
    )
}


## ----glycemic-by-cancer-history, warning=FALSE, message=FALSE, echo=FALSE-----
cancer_plot_data <- a1c_all_a1cdemo_dedup %>%
  dplyr::filter(!is.na(age_group), !is.na(glycemic_category), glycemic_category != "hypoglycemia") %>%
  dplyr::mutate(
    cancer_history = dplyr::case_when(
      cancer == 1 ~ "History of cancer",
      TRUE ~ "No history of cancer"
    ),
    cancer_history = factor(
      cancer_history,
      levels = c("No history of cancer", "History of cancer")
    )
  ) %>%
  dplyr::group_by(age_group, cancer_history, glycemic_category) %>%
  dplyr::summarise(group_count = dplyr::n(), .groups = "drop") %>%
  dplyr::left_join(
    a1c_all_a1cdemo_dedup %>%
      dplyr::filter(!is.na(age_group), !is.na(glycemic_category), glycemic_category != "hypoglycemia") %>%
      dplyr::mutate(
        cancer_history = dplyr::case_when(
          cancer == 1 ~ "History of cancer",
          TRUE ~ "No history of cancer"
        ),
        cancer_history = factor(
          cancer_history,
          levels = c("No history of cancer", "History of cancer")
        )
      ) %>%
      dplyr::group_by(age_group, cancer_history) %>%
      dplyr::summarise(total_donors = dplyr::n(), .groups = "drop"),
    by = c("age_group", "cancer_history")
  ) %>%
  dplyr::mutate(percentage = group_count / total_donors)

cancer_glycemic_tab <- xtabs(
  ~ interaction(age_group, cancer_history) + glycemic_category,
  data = a1c_all_a1cdemo_dedup %>%
    dplyr::filter(!is.na(age_group), !is.na(glycemic_category), glycemic_category != "hypoglycemia") %>%
    dplyr::mutate(
      cancer_history = dplyr::case_when(
        cancer == 1 ~ "History of cancer",
        TRUE ~ "No history of cancer"
      ),
      cancer_history = factor(
        cancer_history,
        levels = c("No history of cancer", "History of cancer")
      ),
      glycemic_category = droplevels(factor(glycemic_category))
    )
)

cancer_glycemic_chisq <- chisq.test(cancer_glycemic_tab)

# knitr::kable(
#   cancer_plot_data %>%
#     dplyr::mutate(
#       percentage = scales::percent(percentage, accuracy = 0.1)
#     ) %>%
#     dplyr::select(
#       `Age group` = age_group,
#       `Cancer history` = cancer_history,
#       `Glycemic category` = glycemic_category,
#       `Donors in category` = group_count,
#       `Total donors in stratum` = total_donors,
#       Percentage = percentage
#     ),
#   caption = "Glycemic status by age group stratified by cancer history."
# )

ggplot(cancer_plot_data, aes(x = age_group, y = percentage)) +
  geom_bar(
    aes(color = glycemic_category, fill = glycemic_category),
    stat = "identity",
    position = position_dodge2(0.8),
    width = 0.7
  ) +
  facet_wrap(~ cancer_history) +
  geom_text(
    aes(
      label = paste0(round(100 * percentage, 1), "%"),
      group = glycemic_category
    ),
    position = position_dodge(0.8),
    vjust = 0,
    hjust = -0.1,
    size = 3.5
  ) +
  coord_flip() +
  scale_fill_brewer(palette = "YlOrRd", direction = 1) +
  scale_color_brewer(palette = "YlOrRd", direction = 1) +
  guides(fill = guide_legend(reverse = TRUE)) +
  guides(color = guide_legend(reverse = TRUE)) +
  scale_y_continuous(labels = scales::percent, limits = c(0, 1.05)) +
  labs(
    title = "Glycemic status of donors by age group stratified by cancer history",
    subtitle = paste(
      "among donors screened in 2025-2026;",
      "chi-square p",
      scales::pvalue(cancer_glycemic_chisq$p.value, accuracy = 0.001)
    ),
    caption = "donors with A1c% below 4.0 are excluded",
    x = "donor age group",
    y = "percentage of donors"
  ) +
  customeTheme +
  theme(
    axis.text.x = element_blank(),
    axis.ticks.x = element_blank(),
    axis.ticks.y = element_blank(),
    axis.text.y = element_text(angle = 0, vjust = 0.5, hjust = 0.5),
    title = element_text(size = 11),
    legend.position = "right",
    rect = element_rect(fill = "transparent"),
    panel.grid.major.y = element_blank(),
    panel.grid.minor.y = element_blank(),
    text = element_text(size = 12),
    panel.background = element_rect(fill = "transparent", colour = NA),
    plot.background = element_rect(fill = "transparent", colour = NA),
    panel.grid.major = element_blank(),
    panel.grid.minor = element_blank()
  )


## ----glycemic-by-combined-hypertension-status, warning=FALSE, message=FALSE, echo=FALSE----
combined_hypertension_plot_data <- a1c_all_a1cdemo_dedup %>%
  dplyr::filter(!is.na(age_group), !is.na(glycemic_category), glycemic_category != "hypoglycemia") %>%
  dplyr::mutate(
    combined_hypertension_status = dplyr::case_when(
      hypertension == 1 |
        BP_category %in% c("Hypertension stage 1", "Hypertension stage 2") ~ "Hypertensive",
      TRUE ~ "Non-hypertensive"
    ),
    combined_hypertension_status = factor(
      combined_hypertension_status,
      levels = c("Non-hypertensive", "Hypertensive")
    )
  ) %>%
  dplyr::group_by(age_group, combined_hypertension_status, glycemic_category) %>%
  dplyr::summarise(group_count = dplyr::n(), .groups = "drop") %>%
  dplyr::left_join(
    a1c_all_a1cdemo_dedup %>%
      dplyr::filter(!is.na(age_group), !is.na(glycemic_category), glycemic_category != "hypoglycemia") %>%
      dplyr::mutate(
        combined_hypertension_status = dplyr::case_when(
          hypertension == 1 |
            BP_category %in% c("Hypertension stage 1", "Hypertension stage 2") ~ "Hypertensive",
          TRUE ~ "Non-hypertensive"
        ),
        combined_hypertension_status = factor(
          combined_hypertension_status,
          levels = c("Non-hypertensive", "Hypertensive")
        )
      ) %>%
      dplyr::group_by(age_group, combined_hypertension_status) %>%
      dplyr::summarise(total_donors = dplyr::n(), .groups = "drop"),
    by = c("age_group", "combined_hypertension_status")
  ) %>%
  dplyr::mutate(percentage = group_count / total_donors)

combined_hypertension_glycemic_tab <- xtabs(
  ~ interaction(age_group, combined_hypertension_status) + glycemic_category,
  data = a1c_all_a1cdemo_dedup %>%
    dplyr::filter(!is.na(age_group), !is.na(glycemic_category), glycemic_category != "hypoglycemia") %>%
    dplyr::mutate(
      combined_hypertension_status = dplyr::case_when(
        hypertension == 1 |
          BP_category %in% c("Hypertension stage 1", "Hypertension stage 2") ~ "Hypertensive",
        TRUE ~ "Non-hypertensive"
      ),
      combined_hypertension_status = factor(
        combined_hypertension_status,
        levels = c("Non-hypertensive", "Hypertensive")
      ),
      glycemic_category = droplevels(factor(glycemic_category))
    )
)

combined_hypertension_glycemic_chisq <- chisq.test(combined_hypertension_glycemic_tab)

# knitr::kable(
#   combined_hypertension_plot_data %>%
#     dplyr::mutate(
#       percentage = scales::percent(percentage, accuracy = 0.1)
#     ) %>%
#     dplyr::select(
#       `Age group` = age_group,
#       `Combined hypertension status` = combined_hypertension_status,
#       `Glycemic category` = glycemic_category,
#       `Donors in category` = group_count,
#       `Total donors in stratum` = total_donors,
#       Percentage = percentage
#     ),
#   caption = "Glycemic status by age group stratified by combined hypertension status."
# )

ggplot(combined_hypertension_plot_data, aes(x = age_group, y = percentage)) +
  geom_bar(
    aes(color = glycemic_category, fill = glycemic_category),
    stat = "identity",
    position = position_dodge2(0.8),
    width = 0.7
  ) +
  facet_wrap(~ combined_hypertension_status) +
  geom_text(
    aes(
      label = paste0(round(100 * percentage, 1), "%"),
      group = glycemic_category
    ),
    position = position_dodge(0.8),
    vjust = 0,
    hjust = -0.1,
    size = 3.5
  ) +
  coord_flip() +
  scale_fill_brewer(palette = "YlOrRd", direction = 1) +
  scale_color_brewer(palette = "YlOrRd", direction = 1) +
  guides(fill = guide_legend(reverse = TRUE)) +
  guides(color = guide_legend(reverse = TRUE)) +
  scale_y_continuous(labels = scales::percent, limits = c(0, 1.05)) +
  labs(
    title = "Glycemic status of donors by age group stratified by combined hypertension status",
    subtitle = paste(
      "among donors screened in 2025-2026;",
      "chi-square p",
      scales::pvalue(combined_hypertension_glycemic_chisq$p.value, accuracy = 0.001)
    ),
    caption = "hypertensive includes donors with hypertension history or BP category Hypertension stage 1 or Hypertension stage 2",
    x = "donor age group",
    y = "percentage of donors"
  ) +
  customeTheme +
  theme(
    axis.text.x = element_blank(),
    axis.ticks.x = element_blank(),
    axis.ticks.y = element_blank(),
    axis.text.y = element_text(angle = 0, vjust = 0.5, hjust = 0.5),
    title = element_text(size = 11),
    legend.position = "right",
    rect = element_rect(fill = "transparent"),
    panel.grid.major.y = element_blank(),
    panel.grid.minor.y = element_blank(),
    text = element_text(size = 12),
    panel.background = element_rect(fill = "transparent", colour = NA),
    plot.background = element_rect(fill = "transparent", colour = NA),
    panel.grid.major = element_blank(),
    panel.grid.minor = element_blank()
  )


## ----donor-kmeans-cluster-prep, warning=FALSE, message=FALSE, include=FALSE----
donor_cluster_results <- tryCatch({
  if (inherits(urban_rural_strat_data, "list") && !is.null(urban_rural_strat_data$error)) {
    stop(urban_rural_strat_data$error)
  }

  donor_cluster_input <- urban_rural_strat_data %>%
    dplyr::mutate(
      donor_type = dplyr::case_when(
        ftd == "Y" ~ "First-time donor",
        ftd == "N" ~ "Repeat donor",
        TRUE ~ NA_character_
      ),
      combined_hypertension_status = dplyr::case_when(
        hypertension == 1 |
          BP_category %in% c("Hypertension stage 1", "Hypertension stage 2") ~ "Hypertensive",
        TRUE ~ "Non-hypertensive"
      ),
      cardiorespiratory_history = dplyr::case_when(
        cardiorespiratory == 1 ~ "History of heart/lung condition",
        cardiorespiratory == 0 ~ "No heart/lung condition",
        TRUE ~ NA_character_
      ),
      statefp = substr(GEOID, 1, 2),
      region = dplyr::case_when(
        statefp %in% c("09", "23", "25", "33", "44", "50", "34", "36", "42") ~ "Northeast",
        statefp %in% c("17", "18", "26", "39", "55", "19", "20", "27", "29", "31", "38", "46") ~ "Midwest",
        statefp %in% c("10", "11", "12", "13", "24", "37", "45", "51", "54", "01", "21", "28", "47", "05", "22", "40", "48") ~ "South",
        statefp %in% c("04", "08", "16", "30", "32", "35", "49", "56", "02", "06", "15", "41", "53") ~ "West",
        TRUE ~ NA_character_
      ),
      age_group = factor(age_group, levels = c("-29", "30-54", "55-")),
      gender = factor(gender, levels = c("F", "M")),
      raceeth = factor(raceeth, levels = c("White", "Black", "Hispanic Origin", "Asian/AIAN/other")),
      donor_type = factor(donor_type, levels = c("First-time donor", "Repeat donor")),
      region = factor(region, levels = c("Northeast", "Midwest", "South", "West")),
      combined_hypertension_status = factor(
        combined_hypertension_status,
        levels = c("Non-hypertensive", "Hypertensive")
      ),
      cardiorespiratory_history = factor(
        cardiorespiratory_history,
        levels = c("No heart/lung condition", "History of heart/lung condition")
      ),
      urban_rural_group = factor(urban_rural_group, levels = c("Urban", "Rural"))
    ) %>%
    dplyr::filter(
      !is.na(age_group),
      !is.na(gender),
      !is.na(raceeth),
      !is.na(donor_type),
      !is.na(region),
      !is.na(combined_hypertension_status),
      !is.na(cardiorespiratory_history),
      !is.na(urban_rural_group),
      !is.na(glycemic_category)
    )

  if (nrow(donor_cluster_input) < 1000) {
    stop("Too few complete-case donors were available for clustering.")
  }

  cluster_formula <- ~ age_group + gender + raceeth + donor_type + region +
    combined_hypertension_status + cardiorespiratory_history + urban_rural_group - 1

  set.seed(42)
  selection_n <- min(10000, nrow(donor_cluster_input))
  selection_idx <- sample.int(nrow(donor_cluster_input), size = selection_n)
  selection_matrix <- stats::model.matrix(cluster_formula, data = donor_cluster_input[selection_idx, , drop = FALSE])

  candidate_k <- 2:6
  selection_stats <- lapply(candidate_k, function(k) {
    km <- stats::kmeans(selection_matrix, centers = k, nstart = 20, iter.max = 100)
    avg_silhouette <- NA_real_

    if (requireNamespace("cluster", quietly = TRUE) && nrow(selection_matrix) <= 5000) {
      sil <- cluster::silhouette(km$cluster, dist(selection_matrix))
      avg_silhouette <- mean(sil[, 3])
    }

    data.frame(
      k = k,
      avg_silhouette = avg_silhouette,
      tot_withinss = km$tot.withinss
    )
  }) %>%
    dplyr::bind_rows()

  best_k <- if (all(is.na(selection_stats$avg_silhouette))) {
    selection_stats$k[which.min(selection_stats$tot_withinss)]
  } else {
    selection_stats$k[which.max(selection_stats$avg_silhouette)]
  }

  training_n <- min(50000, nrow(donor_cluster_input))
  training_idx <- sample.int(nrow(donor_cluster_input), size = training_n)
  training_matrix <- stats::model.matrix(cluster_formula, data = donor_cluster_input[training_idx, , drop = FALSE])

  final_kmeans <- stats::kmeans(
    training_matrix,
    centers = best_k,
    nstart = 30,
    iter.max = 100
  )

  full_matrix <- stats::model.matrix(cluster_formula, data = donor_cluster_input)
  center_sq <- rowSums(final_kmeans$centers^2)
  donor_sq <- rowSums(full_matrix^2)
  distance_matrix <- outer(donor_sq, center_sq, "+") - 2 * (full_matrix %*% t(final_kmeans$centers))
  raw_cluster_id <- max.col(-distance_matrix)

  cluster_order <- donor_cluster_input %>%
    dplyr::mutate(raw_cluster_id = raw_cluster_id) %>%
    dplyr::count(raw_cluster_id, name = "donors") %>%
    dplyr::arrange(dplyr::desc(donors)) %>%
    dplyr::mutate(cluster = paste0("Cluster ", dplyr::row_number()))

  cluster_map <- setNames(cluster_order$cluster, cluster_order$raw_cluster_id)

  donor_cluster_input <- donor_cluster_input %>%
    dplyr::mutate(
      cluster = factor(
        cluster_map[as.character(raw_cluster_id)],
        levels = cluster_order$cluster
      )
    )

  cluster_size_tbl <- donor_cluster_input %>%
    dplyr::count(cluster, name = "donors") %>%
    dplyr::mutate(percentage = donors / sum(donors))

  cluster_profile_vars <- c(
    "age_group", "gender", "raceeth", "donor_type", "region",
    "combined_hypertension_status", "cardiorespiratory_history", "urban_rural_group"
  )

  cluster_profile_labels <- c(
    age_group = "Age group",
    gender = "Gender",
    raceeth = "Race/ethnicity",
    donor_type = "Donor type",
    region = "Region",
    combined_hypertension_status = "Hypertension status",
    cardiorespiratory_history = "Heart/lung condition",
    urban_rural_group = "County type"
  )

  cluster_profile_long <- lapply(cluster_profile_vars, function(var_name) {
    donor_cluster_input %>%
      dplyr::count(cluster, level = .data[[var_name]], name = "n") %>%
      dplyr::group_by(cluster) %>%
      dplyr::mutate(percentage = n / sum(n)) %>%
      dplyr::ungroup() %>%
      dplyr::mutate(
        variable = var_name,
        variable_label = cluster_profile_labels[[var_name]],
        feature = paste0(cluster_profile_labels[[var_name]], ": ", level)
      )
  }) %>%
    dplyr::bind_rows()

  feature_order <- cluster_profile_long %>%
    dplyr::distinct(variable, variable_label, feature, level) %>%
    dplyr::mutate(
      variable = factor(variable, levels = cluster_profile_vars)
    ) %>%
    dplyr::arrange(variable, level) %>%
    dplyr::pull(feature)

  cluster_profile_long <- cluster_profile_long %>%
    dplyr::mutate(
      feature = factor(feature, levels = rev(feature_order))
    )

  cluster_modal_profile <- cluster_profile_long %>%
    dplyr::group_by(cluster, variable_label) %>%
    dplyr::slice_max(order_by = percentage, n = 1, with_ties = FALSE) %>%
    dplyr::ungroup() %>%
    dplyr::mutate(summary = paste0(level, " (", scales::percent(percentage, accuracy = 0.1), ")")) %>%
    dplyr::select(cluster, variable_label, summary) %>%
    tidyr::pivot_wider(names_from = variable_label, values_from = summary)

  cluster_glycemic <- donor_cluster_input %>%
    dplyr::count(cluster, glycemic_category, name = "donors") %>%
    dplyr::group_by(cluster) %>%
    dplyr::mutate(percentage = donors / sum(donors)) %>%
    dplyr::ungroup()

  list(
    donor_cluster_input = donor_cluster_input,
    selection_stats = selection_stats,
    best_k = best_k,
    cluster_size_tbl = cluster_size_tbl,
    cluster_profile_long = cluster_profile_long,
    cluster_modal_profile = cluster_modal_profile,
    cluster_glycemic = cluster_glycemic
  )
}, error = function(e) {
  list(error = conditionMessage(e))
})


## ----donor-kmeans-cluster-output, warning=FALSE, message=FALSE, echo=FALSE, fig.width=11, fig.height=8----
if (!is.null(donor_cluster_results$error)) {
  htmltools::tags$div(
    style = "padding: 12px; border: 1px solid #dddddd; background: #fafafa;",
    htmltools::tags$strong("Donor clustering analysis unavailable."),
    htmltools::tags$div(paste("Reason:", donor_cluster_results$error))
  )
} else {
  htmltools::tagList(
    htmltools::tags$p(
      style = "margin-bottom: 10px;",
      paste0(
        "K-means clustering was performed on ",
        scales::comma(nrow(donor_cluster_results$donor_cluster_input)),
        " complete-case donors using donor demographics, donor type, Census region, combined hypertension status, history of heart/lung conditions, and urban-rural county residence. ",
        "The selected solution retained ",
        donor_cluster_results$best_k,
        " clusters."
      )
    ),
    knitr::kable(
      donor_cluster_results$cluster_size_tbl %>%
        dplyr::mutate(
          donors = scales::comma(donors),
          percentage = scales::percent(percentage, accuracy = 0.1)
        ) %>%
        dplyr::rename(
          Cluster = cluster,
          Donors = donors,
          `Donor share` = percentage
        ),
      caption = "Cluster size for the donor k-means phenotype analysis."
    ),
    knitr::kable(
      donor_cluster_results$cluster_modal_profile,
      caption = "Dominant donor profile within each k-means cluster."
    )
  )
}


## ----donor-kmeans-cluster-plot, warning=FALSE, message=FALSE, echo=FALSE, fig.width=11, fig.height=9----
if (!is.null(donor_cluster_results$error)) {
  plot.new()
  text(0.5, 0.55, "Donor clustering figures unavailable.", cex = 1.1)
  text(0.5, 0.45, paste("Reason:", donor_cluster_results$error), cex = 0.9)
} else {
  cluster_profile_plot <- ggplot(
    donor_cluster_results$cluster_profile_long,
    aes(x = cluster, y = feature, fill = percentage)
  ) +
    geom_tile(color = "white", linewidth = 0.2) +
    geom_text(
      aes(label = ifelse(percentage >= 0.45, scales::percent(percentage, accuracy = 1), "")),
      size = 3,
      color = "black"
    ) +
    scale_fill_gradient(
      low = "#F6E8C3",
      high = "#B35806",
      labels = scales::percent_format(accuracy = 1)
    ) +
    labs(
      title = "Donor phenotype profiles from k-means clustering",
      subtitle = "Each cell shows the within-cluster percentage for a donor characteristic level",
      x = NULL,
      y = NULL,
      fill = "Within-cluster share"
    ) +
    theme_minimal(base_size = 11) +
    theme(
      panel.grid = element_blank(),
      axis.text.x = element_text(face = "bold"),
      plot.title.position = "plot"
    )

  cluster_glycemic_plot <- ggplot(
    donor_cluster_results$cluster_glycemic,
    aes(x = cluster, y = percentage, fill = glycemic_category)
  ) +
    geom_col(width = 0.72) +
    geom_text(
      aes(label = ifelse(percentage >= 0.06, scales::percent(percentage, accuracy = 0.1), "")),
      position = position_stack(vjust = 0.5),
      size = 3.2,
      color = "white",
      fontface = "bold"
    ) +
    scale_fill_manual(
      values = c(
        "Normoglycemia" = "#5B8E7D",
        "Prediabetes" = "#D39C3F",
        "Diabetes" = "#B44B3E"
      )
    ) +
    scale_y_continuous(labels = scales::percent, expand = expansion(mult = c(0, 0.02))) +
    labs(
      title = "Glycemic distribution within donor phenotype clusters",
      subtitle = "Clusters were formed without using glycemic category",
      x = NULL,
      y = "Percentage of donors",
      fill = "Glycemic category"
    ) +
    theme_minimal(base_size = 11) +
    theme(
      panel.grid.minor = element_blank(),
      panel.grid.major.x = element_blank(),
      legend.position = "top",
      plot.title.position = "plot"
    )

  cowplot::plot_grid(
    cluster_profile_plot,
    cluster_glycemic_plot,
    ncol = 1,
    rel_heights = c(2.2, 1.1)
  )
}


## ----diabetes-hotspot-map-prep, warning=FALSE, message=FALSE, include=FALSE----
hotspot_map_data <- tryCatch({
  options(tigris_use_cache = TRUE)

  simplify_leaflet_sf <- function(x, keep_cols, tolerance_m) {
    x <- x %>%
      dplyr::select(dplyr::any_of(keep_cols), geometry)

    if (!nrow(x)) {
      return(sf::st_transform(x, 4326))
    }

    x %>%
      sf::st_transform(2163) %>%
      sf::st_simplify(dTolerance = tolerance_m, preserveTopology = TRUE) %>%
      sf::st_transform(4326)
  }

  diabetes_geo <- a1c_all_a1cdemo_dedup %>%
    dplyr::filter(!is.na(glycemic_category)) %>%
    dplyr::mutate(
      zip5 = stringr::str_pad(
        stringr::str_extract(as.character(zip), "\\d{1,5}"),
        width = 5,
        side = "left",
        pad = "0"
      ),
      diabetes_flag = as.integer(glycemic_category == "Diabetes"),
      age_55_plus = dplyr::case_when(
        !is.na(age) ~ as.numeric(age >= 55),
        age_group == "55-" ~ 1,
        age_group %in% c("-29", "30-54") ~ 0,
        TRUE ~ NA_real_
      )
    ) %>%
    dplyr::filter(!is.na(zip5))

  zip_prefixes <- sort(unique(substr(diabetes_geo$zip5, 1, 3)))

  zcta_sf <- tigris::zctas(
    cb = TRUE,
    starts_with = zip_prefixes,
    year = 2020,
    class = "sf"
  )

  counties_sf <- tigris::counties(
    cb = TRUE,
    year = 2020,
    class = "sf"
  ) %>%
    dplyr::filter(!STATEFP %in% c("60", "66", "69", "72", "78"))

  states_sf <- tigris::states(
    cb = TRUE,
    year = 2020,
    class = "sf"
  ) %>%
    dplyr::filter(!STUSPS %in% c("AS", "GU", "MP", "PR", "VI"))

  state_lookup <- states_sf %>%
    sf::st_drop_geometry() %>%
    dplyr::select(STATEFP, STUSPS)

  format_county_geoid <- function(x) {
    x_chr <- trimws(as.character(x))
    x_chr <- sub("\\.0$", "", x_chr)
    stringr::str_pad(x_chr, width = 5, side = "left", pad = "0")
  }

  urban_rural_county <- readxl::read_excel("urban rural classification CDC.xlsx") %>%
    {
      names(.)[names(.) == "2023 Code"] <- "urban_code"
      .
    } %>%
    dplyr::transmute(
      GEOID = format_county_geoid(Location),
      urban_rural_code = as.character(urban_code),
      urban_rural_group = dplyr::case_when(
        substr(urban_rural_code, 1, 1) %in% c("1", "2", "3", "4") ~ "Urban",
        substr(urban_rural_code, 1, 1) %in% c("5", "6") ~ "Rural",
        TRUE ~ "Unknown"
      )
    )

  svi_county <- readr::read_csv("SVI_2022_US_county.csv", show_col_types = FALSE) %>%
    dplyr::transmute(
      GEOID = format_county_geoid(FIPS),
      svi_overall = as.numeric(RPL_THEMES)
    )

  zip_id_col <- intersect(c("ZCTA5CE20", "GEOID20", "ZCTA5CE10"), names(zcta_sf))[1]

  zcta_to_county <- zcta_sf %>%
    sf::st_transform(5070) %>%
    sf::st_point_on_surface() %>%
    sf::st_transform(sf::st_crs(counties_sf)) %>%
    sf::st_join(counties_sf %>% dplyr::select(GEOID), left = FALSE) %>%
    sf::st_drop_geometry() %>%
    dplyr::transmute(
      zip5 = .data[[zip_id_col]],
      GEOID = .data[["GEOID"]]
    ) %>%
    dplyr::distinct(zip5, .keep_all = TRUE)

  bmi_col <- intersect(c("BMI", "bmi"), names(diabetes_geo))
  bmi_cat_col <- intersect(c("BMI_cat", "bmi_cat"), names(diabetes_geo))

  diabetes_geo <- if (length(bmi_col) > 0) {
    diabetes_geo %>%
      dplyr::mutate(
        obese_flag = as.numeric(
          suppressWarnings(as.numeric(.data[[bmi_col[1]]])) >= 30
        )
      )
  } else if (length(bmi_cat_col) > 0) {
    diabetes_geo %>%
      dplyr::mutate(
        obese_flag = as.numeric(
          stringr::str_to_lower(trimws(as.character(.data[[bmi_cat_col[1]]]))) %in%
            c("obese", "morbid obesity", "morbid obestiy")
        )
      )
  } else {
    diabetes_geo %>%
      dplyr::mutate(obese_flag = NA_real_)
  }

  county_diabetes <- diabetes_geo %>%
    dplyr::left_join(zcta_to_county, by = "zip5") %>%
    dplyr::filter(!is.na(GEOID)) %>%
    dplyr::group_by(GEOID) %>%
    dplyr::summarise(
      n_donors = dplyr::n(),
      diabetes_cases = sum(diabetes_flag, na.rm = TRUE),
      diabetes_prev = 100 * mean(diabetes_flag, na.rm = TRUE),
      median_age = if (all(is.na(age))) NA_real_ else median(age, na.rm = TRUE),
      pct_age_55_plus = if (all(is.na(age_55_plus))) NA_real_ else 100 * mean(age_55_plus, na.rm = TRUE),
      pct_obese = 100 * mean(obese_flag, na.rm = TRUE),
      .groups = "drop"
    )

  hotspot_input <- counties_sf %>%
    dplyr::left_join(county_diabetes, by = "GEOID") %>%
    dplyr::filter(!is.na(diabetes_prev), n_donors >= 30)

  county_nb <- spdep::poly2nb(hotspot_input, queen = TRUE)
  county_lw <- spdep::nb2listw(county_nb, style = "W", zero.policy = TRUE)
  hotspot_input$gi_star <- as.numeric(
    spdep::localG(hotspot_input$diabetes_prev, county_lw, zero.policy = TRUE)
  )

  hotspot_class <- hotspot_input %>%
    sf::st_drop_geometry() %>%
    dplyr::transmute(
      GEOID,
      hotspot = dplyr::case_when(
        gi_star >= 1.96 ~ "Hot spot",
        gi_star <= -1.96 ~ "Cold spot",
        TRUE ~ "Not significant"
      )
    )

  county_map <- counties_sf %>%
    {
      if ("STUSPS" %in% names(.)) . else dplyr::left_join(., state_lookup, by = "STATEFP")
    } %>%
    dplyr::left_join(county_diabetes, by = "GEOID") %>%
    dplyr::left_join(urban_rural_county, by = "GEOID") %>%
    dplyr::left_join(svi_county, by = "GEOID") %>%
    dplyr::left_join(hotspot_class, by = "GEOID") %>%
    dplyr::mutate(
      hotspot = dplyr::case_when(
        is.na(n_donors) ~ "No donor data",
        n_donors < 30 ~ "Insufficient donor count",
        TRUE ~ hotspot
      ),
      hotspot = factor(
        hotspot,
        levels = c(
          "Hot spot",
          "Cold spot",
          "Not significant",
          "Insufficient donor count",
          "No donor data"
        )
      ),
      county_label = paste0(.data[["NAME"]], ", ", .data[["STUSPS"]]),
      popup_html = dplyr::case_when(
        is.na(n_donors) ~ paste0(
          "<strong>", county_label, "</strong><br/>No donor data available."
        ),
        TRUE ~ paste0(
          "<strong>", county_label, "</strong>",
          "<br/>Diabetic donors: ", scales::percent(diabetes_prev / 100, accuracy = 0.1),
          "<br/>Median donor age: ", ifelse(is.na(median_age), "Missing", scales::number(median_age, accuracy = 0.1)),
          "<br/>Donors age 55+: ", ifelse(is.na(pct_age_55_plus), "Missing", scales::percent(pct_age_55_plus / 100, accuracy = 0.1)),
          "<br/>Urban-rural class: ", ifelse(is.na(urban_rural_code), "Missing", urban_rural_code),
          "<br/>Overall SVI percentile: ", ifelse(is.na(svi_overall), "Missing", scales::number(svi_overall, accuracy = 0.01)),
          "<br/>Donor count: ", scales::comma(n_donors),
          "<br/>County pattern: ", as.character(hotspot)
        )
      ),
      hover_label = dplyr::case_when(
        is.na(n_donors) ~ paste0(
          "<strong>", county_label, "</strong><br/>No donor data"
        ),
        TRUE ~ paste0(
          "<strong>", county_label, "</strong>",
          "<br/>Diabetic donors: ", scales::percent(diabetes_prev / 100, accuracy = 0.1)
        )
      ),
      fill_color = dplyr::case_when(
        hotspot == "Hot spot" ~ "#B5413C",
        hotspot == "Cold spot" ~ "#4E79A7",
        hotspot == "Not significant" ~ "#D9D9D9",
        hotspot == "Insufficient donor count" ~ "#F2E6C9",
        hotspot == "No donor data" ~ "#F7F7F7",
        TRUE ~ "#F7F7F7"
      )
    )

  hotspot_county_map <- county_map %>%
    dplyr::filter(!is.na(n_donors), n_donors >= 30)

  hotspot_county_map_plot <- tigris::shift_geometry(hotspot_county_map, position = "outside")
  county_map_plot <- tigris::shift_geometry(county_map, position = "outside")
  states_plot <- tigris::shift_geometry(states_sf, position = "outside")
  hotspot_county_map_leaflet <- simplify_leaflet_sf(
    hotspot_county_map_plot,
    keep_cols = c("fill_color", "hover_label", "popup_html", "diabetes_prev"),
    tolerance_m = 2500
  )
  county_map_leaflet <- simplify_leaflet_sf(
    county_map_plot,
    keep_cols = c(
      "GEOID",
      "county_label",
      "diabetes_prev",
      "urban_rural_code",
      "urban_rural_group",
      "svi_overall",
      "popup_html"
    ),
    tolerance_m = 2500
  )
  states_leaflet <- simplify_leaflet_sf(
    states_plot,
    keep_cols = c("STUSPS"),
    tolerance_m = 5000
  )

  list(
    county_map = county_map,
    hotspot_county_map = hotspot_county_map,
    county_map_plot = county_map_plot,
    hotspot_county_map_plot = hotspot_county_map_plot,
    county_map_leaflet = county_map_leaflet,
    hotspot_county_map_leaflet = hotspot_county_map_leaflet,
    states_leaflet = states_leaflet,
    states_plot = states_plot,
    zcta_to_county = zcta_to_county,
    urban_rural_county = urban_rural_county,
    county_analysis = county_map %>%
      sf::st_drop_geometry() %>%
      dplyr::filter(!is.na(diabetes_prev), !is.na(svi_overall), n_donors >= 30)
  )
}, error = function(e) {
  list(error = conditionMessage(e))
})


## ----diabetes-hotspot-map, echo=FALSE, warning=FALSE, message=FALSE, fig.width=12, fig.height=7----
if (!is.null(hotspot_map_data$error)) {
  plot.new()
  text(
    0.5, 0.58,
    "The diabetes hotspot map could not be rendered in this session.",
    cex = 1.1
  )
  text(
    0.5, 0.48,
    paste("Reason:", hotspot_map_data$error),
    cex = 0.9
  )
} else {
  ggplot(hotspot_map_data$hotspot_county_map_plot) +
    geom_sf(aes(fill = hotspot), color = NA) +
    geom_sf(data = hotspot_map_data$states_plot, fill = NA, color = "white", linewidth = 0.15) +
    scale_fill_manual(
      values = c(
        "Hot spot" = "#B5413C",
        "Cold spot" = "#4E79A7",
        "Not significant" = "#D9D9D9"
      ),
      drop = FALSE
    ) +
    labs(
      title = "County hot spots of donor diabetes prevalence",
      subtitle = "Latest HbA1c per donor; county prevalence built from donor ZIP and evaluated for local clustering in counties with at least 30 donors",
      fill = "County pattern",
      caption = "CDC reference: U.S. Diabetes Surveillance System spotlight map. Only counties with at least 30 donors are shown."
    ) +
    theme_void(base_size = 12) +
    theme(
      legend.position = "right",
      plot.title.position = "plot",
      plot.caption.position = "plot"
    )
}


## ----diabetes-hotspot-map-interactive, echo=FALSE, warning=FALSE, message=FALSE----
if (!is.null(hotspot_map_data$error)) {
  htmltools::tags$div(
    style = "padding: 12px; border: 1px solid #dddddd; background: #fafafa;",
    htmltools::tags$strong("Interactive hotspot map unavailable."),
    htmltools::tags$div(paste("Reason:", hotspot_map_data$error))
  )
} else {
  hotspot_colors <- c(
    "Hot spot" = "#B5413C",
    "Cold spot" = "#4E79A7",
    "Not significant" = "#D9D9D9"
  )

  county_map_leaflet <- hotspot_map_data$hotspot_county_map_leaflet
  states_leaflet <- hotspot_map_data$states_leaflet
  diabetes_pal <- leaflet::colorNumeric(
    palette = c("#fff5eb", "#fdd0a2", "#f16913", "#b30000"),
    domain = county_map_leaflet$diabetes_prev,
    na.color = "#F7F7F7"
  )
  county_map_leaflet$diabetes_fill <- diabetes_pal(county_map_leaflet$diabetes_prev)

  leaflet::leaflet(
    county_map_leaflet,
    width = "100%",
    height = 650,
    options = leaflet::leafletOptions(minZoom = 3, zoomSnap = 0.25)
  ) %>%
    leaflet::addProviderTiles("CartoDB.Positron") %>%
    leaflet::addPolygons(
      group = "County hotspot class",
      fillColor = ~ fill_color,
      fillOpacity = 0.85,
      color = "white",
      weight = 0.4,
      smoothFactor = 0.2,
      label = ~ lapply(hover_label, htmltools::HTML),
      labelOptions = leaflet::labelOptions(
        direction = "auto",
        style = list(
          "font-weight" = "normal",
          "padding" = "6px 8px"
        ),
        textsize = "13px"
      ),
      popup = ~ popup_html,
      highlightOptions = leaflet::highlightOptions(
        weight = 1.2,
        color = "#2F2F2F",
        fillOpacity = 0.95,
        bringToFront = TRUE
      )
    ) %>%
    leaflet::addPolygons(
      group = "Donor diabetes %",
      fillColor = ~ diabetes_fill,
      fillOpacity = 0.85,
      color = "white",
      weight = 0.4,
      smoothFactor = 0.2,
      label = ~ lapply(hover_label, htmltools::HTML),
      labelOptions = leaflet::labelOptions(
        direction = "auto",
        style = list(
          "font-weight" = "normal",
          "padding" = "6px 8px"
        ),
        textsize = "13px"
      ),
      popup = ~ popup_html,
      highlightOptions = leaflet::highlightOptions(
        weight = 1.2,
        color = "#2F2F2F",
        fillOpacity = 0.95,
        bringToFront = TRUE
      )
    ) %>%
    leaflet::addPolylines(
      data = states_leaflet,
      color = "#FFFFFF",
      weight = 0.5,
      opacity = 0.9
    ) %>%
    leaflet::addLegend(
      position = "bottomright",
      colors = hotspot_colors,
      labels = names(hotspot_colors),
      opacity = 0.9,
      title = "County hotspot class",
      group = "County hotspot class"
    ) %>%
    leaflet::addLegend(
      position = "bottomright",
      pal = diabetes_pal,
      values = ~ diabetes_prev,
      opacity = 0.9,
      title = "Donor diabetes %",
      group = "Donor diabetes %",
      labFormat = leaflet::labelFormat(suffix = "%")
    ) %>%
    leaflet::addLayersControl(
      baseGroups = c("County hotspot class", "Donor diabetes %"),
      options = leaflet::layersControlOptions(collapsed = FALSE)
    ) %>%
    leaflet::hideGroup("Donor diabetes %") %>%
    leaflet::setView(lng = -96, lat = 38, zoom = 4)
}


## ----county-svi-swipe-map, echo=FALSE, warning=FALSE, message=FALSE-----------
if (!is.null(hotspot_map_data$error)) {
  htmltools::tags$div(
    style = "padding: 12px; border: 1px solid #dddddd; background: #fafafa;",
    htmltools::tags$strong("Swipe map unavailable."),
    htmltools::tags$div(paste("Reason:", hotspot_map_data$error))
  )
} else {
  swipe_county_map <- hotspot_map_data$county_map_leaflet
  swipe_states <- hotspot_map_data$states_leaflet

  diabetes_values <- swipe_county_map$diabetes_prev[!is.na(swipe_county_map$diabetes_prev)]
  if (length(diabetes_values) > 1) {
    diabetes_bins <- unique(as.numeric(stats::quantile(diabetes_values, probs = seq(0, 1, length.out = 6), na.rm = TRUE)))
  } else {
    diabetes_bins <- c(0, 5, 7.5, 10, 12.5, 20)
  }
  if (length(diabetes_bins) < 3) {
    diabetes_bins <- c(0, 5, 7.5, 10, 12.5, 20)
  }

  urban_diabetes_pal <- leaflet::colorBin(
    palette = c("#efedf5", "#cbc9e2", "#9e9ac8", "#756bb1", "#54278f"),
    domain = diabetes_values,
    bins = diabetes_bins,
    na.color = "#efe9df"
  )

  rural_diabetes_pal <- leaflet::colorBin(
    palette = c("#edf8fb", "#b2e2e2", "#66c2a4", "#2ca25f", "#006d2c"),
    domain = diabetes_values,
    bins = diabetes_bins,
    na.color = "#efe9df"
  )

  svi_pal <- leaflet::colorBin(
    palette = c("#fff7bc", "#fec44f", "#f03b20", "#bd0026"),
    domain = swipe_county_map$svi_overall,
    bins = c(0, 0.25, 0.50, 0.75, 1.00),
    na.color = "#efe9df"
  )

  swipe_county_map <- swipe_county_map %>%
    dplyr::mutate(
      diabetes_swipe_fill = dplyr::case_when(
        urban_rural_group == "Urban" ~ urban_diabetes_pal(diabetes_prev),
        urban_rural_group == "Rural" ~ rural_diabetes_pal(diabetes_prev),
        TRUE ~ "#efe9df"
      ),
      svi_swipe_fill = svi_pal(svi_overall),
      diabetes_layer_id = paste0("diabetes_", GEOID),
      svi_layer_id = paste0("svi_", GEOID),
      swipe_label = paste0(
        "<strong>", county_label, "</strong>",
        "<br/>Donor diabetes: ", ifelse(is.na(diabetes_prev), "Missing", scales::percent(diabetes_prev / 100, accuracy = 0.1)),
        "<br/>Urban-rural: ", ifelse(is.na(urban_rural_code), "Missing", urban_rural_code),
        "<br/>Overall SVI: ", ifelse(is.na(svi_overall), "Missing", scales::number(svi_overall, accuracy = 0.01))
      )
    )

  swipe_leaflet_widget <- leaflet::leaflet(
    width = "100%",
    height = 680,
    options = leaflet::leafletOptions(zoomSnap = 0.25, minZoom = 2)
  ) %>%
    leaflet::addMapPane("left", zIndex = 410) %>%
    leaflet::addMapPane("right", zIndex = 420) %>%
    leaflet::addProviderTiles(leaflet::providers$CartoDB.Positron) %>%
    leaflet::addPolygons(
      data = swipe_county_map,
      layerId = ~ diabetes_layer_id,
      fillColor = ~ diabetes_swipe_fill,
      fillOpacity = 0.92,
      color = "#ffffff",
      weight = 0.20,
      smoothFactor = 0,
      popup = ~ popup_html,
      label = ~ lapply(swipe_label, htmltools::HTML),
      options = leaflet::pathOptions(pane = "left"),
      highlightOptions = leaflet::highlightOptions(
        weight = 1.2,
        color = "#2f2a24",
        fillOpacity = 1,
        bringToFront = TRUE
      )
    ) %>%
    leaflet::addPolygons(
      data = swipe_county_map,
      layerId = ~ svi_layer_id,
      fillColor = ~ svi_swipe_fill,
      fillOpacity = 0.92,
      color = "#ffffff",
      weight = 0.20,
      smoothFactor = 0,
      popup = ~ popup_html,
      label = ~ lapply(swipe_label, htmltools::HTML),
      options = leaflet::pathOptions(pane = "right"),
      highlightOptions = leaflet::highlightOptions(
        weight = 1.2,
        color = "#2f2a24",
        fillOpacity = 1,
        bringToFront = TRUE
      )
    ) %>%
    leaflet::addPolylines(
      data = swipe_states,
      color = "#8d857d",
      weight = 0.6,
      opacity = 1,
      options = leaflet::pathOptions(interactive = FALSE)
    ) %>%
    leaflet::addControl(
      html = as.character(
        htmltools::tags$div(
          style = paste(
            "background: rgba(255,255,255,0.92);",
            "padding: 8px 10px;",
            "border-radius: 8px;",
            "font-size: 12px;",
            "line-height: 1.35;"
          ),
          htmltools::tags$div(
            style = "font-weight: 600; margin-bottom: 4px;",
            "Swipe comparison"
          ),
          htmltools::tags$div("Left: donor diabetes prevalence by urban-rural county class"),
          htmltools::tags$div("Right: county overall SVI percentile")
        )
      ),
      position = "topright"
    ) %>%
    leaflet::addLegend(
      position = "bottomleft",
      pal = urban_diabetes_pal,
      values = diabetes_values,
      title = "URBAN donor diabetes %",
      opacity = 1,
      labFormat = leaflet::labelFormat(suffix = "%")
    ) %>%
    leaflet::addLegend(
      position = "bottomleft",
      pal = rural_diabetes_pal,
      values = diabetes_values,
      title = "RURAL donor diabetes %",
      opacity = 1,
      labFormat = leaflet::labelFormat(suffix = "%")
    ) %>%
    leaflet::addLegend(
      position = "bottomright",
      pal = svi_pal,
      values = swipe_county_map$svi_overall,
      title = "Overall SVI percentile",
      opacity = 1,
      labFormat = leaflet::labelFormat(digits = 2)
    ) %>%
    leaflet::setView(lng = -96, lat = 37, zoom = 4) %>%
    htmlwidgets::onRender(
      "
      function(el, x) {
        var map = this;
        var leftPane = map.getPane('left');
        var rightPane = map.getPane('right');
        if (!leftPane || !rightPane) return;

        var existing = el.querySelector('.custom-swipe-wrap');
        if (existing) existing.remove();

        var wrap = document.createElement('div');
        wrap.className = 'custom-swipe-wrap';
        wrap.style.position = 'absolute';
        wrap.style.top = '0';
        wrap.style.left = '0';
        wrap.style.right = '0';
        wrap.style.bottom = '0';
        wrap.style.pointerEvents = 'none';
        wrap.style.zIndex = '1000';

        var divider = document.createElement('div');
        divider.style.position = 'absolute';
        divider.style.top = '0';
        divider.style.bottom = '0';
        divider.style.width = '4px';
        divider.style.marginLeft = '-2px';
        divider.style.backgroundColor = '#ffffff';
        divider.style.boxShadow = '0 0 0 1px rgba(0,0,0,0.12)';
        divider.style.pointerEvents = 'none';
        wrap.appendChild(divider);

        var handle = document.createElement('div');
        handle.style.position = 'absolute';
        handle.style.top = '50%';
        handle.style.width = '34px';
        handle.style.height = '34px';
        handle.style.marginLeft = '-17px';
        handle.style.marginTop = '-17px';
        handle.style.borderRadius = '50%';
        handle.style.backgroundColor = 'rgba(255,255,255,0.95)';
        handle.style.boxShadow = '0 1px 6px rgba(0,0,0,0.25)';
        handle.style.border = '1px solid rgba(0,0,0,0.18)';
        handle.style.pointerEvents = 'auto';
        handle.style.cursor = 'ew-resize';
        handle.style.display = 'flex';
        handle.style.alignItems = 'center';
        handle.style.justifyContent = 'center';
        handle.style.color = '#6b7280';
        handle.style.fontSize = '18px';
        handle.style.fontWeight = '700';
        handle.innerHTML = '&#9474;&#9474;';
        wrap.appendChild(handle);

        el.appendChild(wrap);

        var swipeValue = 0.5;
        var dragging = false;

        function applyClip(value) {
          swipeValue = Math.max(0, Math.min(1, value));
          var size = map.getSize();
          var dividerX = size.x * swipeValue;
          divider.style.left = dividerX + 'px';
          handle.style.left = dividerX + 'px';
          leftPane.style.clip = 'rect(' + [0, dividerX, size.y, 0].join('px,') + 'px)';
          rightPane.style.clip = 'rect(' + [0, size.x, size.y, dividerX].join('px,') + 'px)';
        }

        function valueFromClientX(clientX) {
          var rect = el.getBoundingClientRect();
          return (clientX - rect.left) / rect.width;
        }

        function startDrag(e) {
          dragging = true;
          if (map.dragging && map.dragging.enabled()) map.dragging.disable();
          e.preventDefault();
        }

        function moveDrag(e) {
          if (!dragging) return;
          var clientX = e.touches ? e.touches[0].clientX : e.clientX;
          applyClip(valueFromClientX(clientX));
        }

        function endDrag() {
          if (!dragging) return;
          dragging = false;
          if (map.dragging && !map.dragging.enabled()) map.dragging.enable();
        }

        handle.addEventListener('mousedown', startDrag);
        handle.addEventListener('touchstart', startDrag, { passive: false });
        document.addEventListener('mousemove', moveDrag);
        document.addEventListener('touchmove', moveDrag, { passive: false });
        document.addEventListener('mouseup', endDrag);
        document.addEventListener('touchend', endDrag);

        map.on('move zoom resize', function() {
          applyClip(swipeValue);
        });

        applyClip(0.5);
      }
      "
    )

  swipe_leaflet_widget
}


## ----county-association-prep, include=FALSE-----------------------------------
if (is.null(hotspot_map_data$error)) {
  county_assoc <- hotspot_map_data$county_analysis

  svi_assoc_test <- cor.test(
    county_assoc$diabetes_prev,
    county_assoc$svi_overall,
    method = "spearman",
    exact = FALSE
  )

  urban_rural_assoc <- county_assoc %>%
    dplyr::filter(urban_rural_group %in% c("Urban", "Rural"))

  urban_rural_summary <- urban_rural_assoc %>%
    dplyr::group_by(urban_rural_group) %>%
    dplyr::summarise(
      counties = dplyr::n(),
      median_diabetes_pct = median(diabetes_prev, na.rm = TRUE),
      mean_diabetes_pct = mean(diabetes_prev, na.rm = TRUE),
      weighted_mean_diabetes_pct = weighted.mean(diabetes_prev, n_donors, na.rm = TRUE),
      donor_count = sum(n_donors, na.rm = TRUE),
      .groups = "drop"
    )

  urban_rural_ttest <- t.test(
    diabetes_prev ~ urban_rural_group,
    data = urban_rural_assoc
  )
}


## ----county-svi-scatterplot, echo=FALSE, warning=FALSE, message=FALSE, fig.width=8.5, fig.height=6----
if (is.null(hotspot_map_data$error)) {
  county_assoc_plot_base <- ggplot(
    county_assoc,
    aes(x = svi_overall, y = diabetes_prev)
  ) +
    geom_smooth(
      method = "lm",
      se = TRUE,
      color = "#17324D",
      fill = "#9FB3C8",
      linewidth = 0.9
    ) +
    scale_color_gradientn(
      colours = c("#F6E8C3", "#D8B365", "#B35806", "#7F0000"),
      na.value = "#BDBDBD",
      labels = function(x) paste0(scales::number(x, accuracy = 0.1), "%")
    ) +
    scale_size_continuous(range = c(1.2, 6), labels = scales::comma) +
    scale_x_continuous(labels = scales::number_format(accuracy = 0.01)) +
    scale_y_continuous(labels = function(x) paste0(x, "%")) +
    labs(
      title = "County donor diabetes prevalence vs. county SVI",
      subtitle = paste0(
        "Spearman rho = ",
        round(unname(svi_assoc_test$estimate), 3),
        ", p = ",
        scales::pvalue(svi_assoc_test$p.value, accuracy = 0.001)
      ),
      x = "Overall SVI percentile",
      y = "Diabetic donors (%)",
      size = "Donors per county",
      color = "Obese donors (%)"
    ) +
    theme_minimal(base_size = 12) +
    theme(
      panel.grid.minor = element_blank(),
      plot.title.position = "plot",
      legend.position = "right"
    )

  county_assoc_plot_base +
    geom_point(
      aes(size = n_donors, color = pct_obese),
      alpha = 0.55
    )
}


## ----urban-rural-diabetes-summary, echo=FALSE, warning=FALSE, message=FALSE----
if (is.null(hotspot_map_data$error)) {
  rural_mean <- urban_rural_summary %>%
    dplyr::filter(urban_rural_group == "Rural") %>%
    dplyr::pull(mean_diabetes_pct)

  urban_mean <- urban_rural_summary %>%
    dplyr::filter(urban_rural_group == "Urban") %>%
    dplyr::pull(mean_diabetes_pct)

  cat(
    paste0(
      "Across counties with at least 30 donors, the mean diabetic-donor percentage was ",
      scales::number(urban_mean, accuracy = 0.01),
      "% in urban counties and ",
      scales::number(rural_mean, accuracy = 0.01),
      "% in rural counties. ",
      "The unweighted county-level difference (rural minus urban) was ",
      scales::number(rural_mean - urban_mean, accuracy = 0.01),
      " percentage points (t-test p = ",
      scales::pvalue(urban_rural_ttest$p.value, accuracy = 0.001),
      ")."
    )
  )

  knitr::kable(
    urban_rural_summary %>%
      dplyr::mutate(
        median_diabetes_pct = scales::number(median_diabetes_pct, accuracy = 0.01),
        mean_diabetes_pct = scales::number(mean_diabetes_pct, accuracy = 0.01),
        weighted_mean_diabetes_pct = scales::number(weighted_mean_diabetes_pct, accuracy = 0.01),
        donor_count = scales::comma(donor_count)
      ),
    col.names = c(
      "County type", "Counties", "Median diabetic donors (%)",
      "Mean diabetic donors (%)", "Weighted mean diabetic donors (%)", "Donors"
    )
  )
}


## ----donor-retention-sankey, echo=FALSE, warning=FALSE, message=FALSE, echo=FALSE----
if (requireNamespace("networkD3", quietly = TRUE)) {
  total_2025_donors <- nrow(ids_2025)
  returned_pct <- scales::percent(retention_share_2025_to_2026, accuracy = 0.1)
  not_yet_returned_pct <- scales::percent(
    1 - retention_share_2025_to_2026,
    accuracy = 0.1
  )

  retention_nodes <- tibble::tibble(
    name = c(
      paste0("Total donors 2025 (", scales::comma(total_2025_donors), "; 100%)"),
      paste0(
        "Rescreened in Mar 2026 (",
        scales::comma(repeat_donors_2025_2026),
        "; ",
        returned_pct,
        ")"
      ),
      paste0(
        "Not yet rescreened (",
        scales::comma(total_2025_donors - repeat_donors_2025_2026),
        "; ",
        not_yet_returned_pct,
        ")"
      )
    ),
    group = c("cohort", "returned", "not_returned")
  )

  retention_links <- retention_flow %>%
    dplyr::mutate(
      source = 0L,
      target = dplyr::if_else(
        as.character(`Mar 2026 status`) == "Rescreened in Mar 2026",
        1L,
        2L
      ),
      value = donors,
      group = dplyr::if_else(
        as.character(`Mar 2026 status`) == "Rescreened in Mar 2026",
        "returned",
        "not_returned"
      )
    ) %>%
    dplyr::select(source, target, value, group)

  retention_sankey <- networkD3::sankeyNetwork(
    Links = retention_links,
    Nodes = retention_nodes,
    Source = "source",
    Target = "target",
    Value = "value",
    NodeID = "name",
    NodeGroup = "group",
    LinkGroup = "group",
    colourScale = htmlwidgets::JS(
      "d3.scaleOrdinal()
        .domain(['cohort', 'returned', 'not_returned'])
        .range(['#D8C3A5', '#1F6C73', '#B86B52'])"
    ),
    fontSize = 14,
    fontFamily = "Arial",
    nodeWidth = 34,
    nodePadding = 28,
    iterations = 0,
    sinksRight = TRUE,
    width = 900,
    height = 420
  )

  htmltools::tagList(
    htmltools::tags$div(
      style = "margin-bottom: 10px;",
      htmltools::tags$h4(
        style = "margin-bottom: 4px;",
        "Retention from 2025 donors into Mar 2026"
      ),
      htmltools::tags$p(
        style = "margin-top: 0; color: #4B5563;",
        paste0(
          scales::percent(retention_share_2025_to_2026, accuracy = 0.1),
          " of donors screened in 2025 were rescreened in Mar 2026."
        )
      )
    ),
    retention_sankey
  )
} else {
  plot.new()
  text(
    0.5, 0.55,
    "Install the networkD3 package to render the retention Sankey plot.",
    cex = 1.1
  )
}


## ----repeat-a1c-pmatrix, echo=FALSE, warning=FALSE, message=FALSE-------------
repeat_p_estimate <- if (is.list(repeat_pmatrix_1y) && !is.null(repeat_pmatrix_1y$estimate)) {
  repeat_pmatrix_1y$estimate
} else {
  repeat_pmatrix_1y
}

repeat_p_raw_mat <- as.matrix(repeat_p_estimate)
repeat_p_mat <- matrix(as.numeric(repeat_p_raw_mat), nrow = 3, ncol = 3)
rownames(repeat_p_mat) <- repeat_state_labels
colnames(repeat_p_mat) <- repeat_state_labels

repeat_p_tbl <- tibble::tibble(
  `Current state` = rownames(repeat_p_mat),
  Normoglycemia = round(repeat_p_mat[, 1], 3),
  Prediabetes = round(repeat_p_mat[, 2], 3),
  Diabetes = round(repeat_p_mat[, 3], 3)
)
knitr::kable(
  repeat_p_tbl,
  caption = "Estimated one-year transition probability matrix from the repeat-donor disease progression model."
)


## ----repeat-a1c-pmatrix-sankey, echo=FALSE, warning=FALSE, message=FALSE------
repeat_observed_transition_shares <- a1c_repeat_subset %>%
  dplyr::arrange(id, .data[["Donation Date"]], visit_num) %>%
  dplyr::group_by(id) %>%
  dplyr::transmute(
    from_state = as.character(glycemic_category),
    to_state = dplyr::lead(as.character(glycemic_category))
  ) %>%
  dplyr::ungroup() %>%
  dplyr::filter(!is.na(from_state), !is.na(to_state)) %>%
  dplyr::count(from_state, to_state, name = "n_transitions") %>%
  tidyr::complete(
    from_state = repeat_state_labels,
    to_state = repeat_state_labels,
    fill = list(n_transitions = 0)
  ) %>%
  dplyr::group_by(from_state) %>%
  dplyr::mutate(
    from_total = sum(n_transitions),
    observed_share = dplyr::if_else(from_total > 0, n_transitions / from_total, NA_real_),
    observed_pct_label = scales::percent(observed_share, accuracy = 0.1)
  ) %>%
  dplyr::ungroup() %>%
  dplyr::select(from_state, to_state, n_transitions, observed_share, observed_pct_label)

repeat_probability_sankey <- tibble::as_tibble(
  repeat_p_mat,
  rownames = "from_state"
) %>%
  tidyr::pivot_longer(
    cols = -from_state,
    names_to = "to_state",
    values_to = "probability"
  ) %>%
  dplyr::left_join(
    repeat_observed_transition_shares,
    by = c("from_state", "to_state")
  ) %>%
  dplyr::mutate(
    from = dplyr::case_when(
      from_state == "Normoglycemia" ~ "current_normoglycemia",
      from_state == "Prediabetes" ~ "current_prediabetes",
      from_state == "Diabetes" ~ "current_diabetes"
    ),
    to = dplyr::case_when(
      to_state == "Normoglycemia" ~ "next_normoglycemia",
      to_state == "Prediabetes" ~ "next_prediabetes",
      to_state == "Diabetes" ~ "next_diabetes"
    ),
    model_pct_label = scales::percent(probability, accuracy = 0.1),
    link_group = dplyr::case_when(
      to_state == "Diabetes" ~ "to_diabetes",
      to_state == "Normoglycemia" ~ "to_normoglycemia",
      TRUE ~ "to_prediabetes"
    ),
    source_rank = match(from_state, repeat_state_labels),
    target_rank = match(to_state, repeat_state_labels),
    label_pos = dplyr::case_when(
      from_state == "Prediabetes" & to_state == "Normoglycemia" ~ 0.68,
      target_rank < source_rank ~ 0.34,
      target_rank > source_rank ~ 0.86,
      TRUE ~ 0.58
    ),
    label_dy = dplyr::case_when(
      target_rank == 1L ~ -6,
      target_rank == 2L ~ 12,
      TRUE ~ 24
    )
  ) %>%
  dplyr::select(
    from_state, to_state, from, to, probability, model_pct_label,
    n_transitions, observed_pct_label, label_pos, label_dy, link_group
  )

repeat_probability_sankey_nodes <- tibble::tibble(
  id = c(
    "current_normoglycemia", "current_prediabetes", "current_diabetes",
    "next_normoglycemia", "next_prediabetes", "next_diabetes"
  ),
  name = c(
    "Current: Normoglycemia", "Current: Prediabetes", "Current: Diabetes",
    "One year: Normoglycemia", "One year: Prediabetes", "One year: Diabetes"
  )
)

repeat_observed_count_sankey_nodes <- tibble::tibble(
  id = c(
    "current_normoglycemia", "current_prediabetes", "current_diabetes",
    "observed_normoglycemia", "observed_prediabetes", "observed_diabetes"
  ),
  name = c(
    "Current: Normoglycemia", "Current: Prediabetes", "Current: Diabetes",
    "Next observed: Normoglycemia", "Next observed: Prediabetes", "Next observed: Diabetes"
  )
)

build_repeat_sankey_widget <- function(
  links_df,
  nodes_df,
  value_col,
  band_label_col,
  units,
  node_width = 28,
  width = 900,
  height = 420
) {
  link_data <- links_df %>%
    dplyr::mutate(
      source = match(from, nodes_df$id) - 1L,
      target = match(to, nodes_df$id) - 1L,
      value = .data[[value_col]]
    ) %>%
    dplyr::select(
      source, target, value, model_pct_label, observed_pct_label,
      n_transitions, link_group, label_pos, label_dy,
      band_label = dplyr::all_of(band_label_col)
    )

  sankey_plot <- networkD3::sankeyNetwork(
    Links = link_data,
    Nodes = nodes_df,
    Source = "source",
    Target = "target",
    Value = "value",
    NodeID = "name",
    units = units,
    fontSize = 13,
    nodeWidth = node_width,
    width = width,
    height = height,
    sinksRight = TRUE
  )

  htmlwidgets::onRender(
    sankey_plot,
    '
    function(el, x, data) {
      setTimeout(function() {
        var svg = d3.select(el).select("svg");
        var links = svg.selectAll(".link");
        var lookup = {};

        links
          .style("stroke", function(d) {
            return d.target.name.indexOf("Diabetes") !== -1
              ? "rgba(180, 75, 62, 0.55)"
              : d.target.name.indexOf("Normoglycemia") !== -1
                ? "rgba(91, 142, 125, 0.48)"
                : "rgba(211, 156, 63, 0.55)";
          })
          .style("stroke-opacity", 1);

        data.keys.forEach(function(key, idx) {
          lookup[key] = {
            band_label: data.band_labels[idx],
            observed_pct: data.observed_pcts[idx],
            model_pct: data.model_pcts[idx],
            n: data.counts[idx],
            pos: data.pos[idx],
            dy: data.dy[idx],
            tooltip: data.tooltips[idx]
          };
        });

        links.each(function(d) {
          var key = d.source.name + "|" + d.target.name;
          var meta = lookup[key] || { tooltip: "" };

          d3.select(this).select("title").remove();
          d3.select(this)
            .append("title")
            .text(meta.tooltip);
        });

        svg.selectAll(".repeat-pmatrix-label-layer").remove();

        var labelLayer = svg.append("g")
          .attr("class", "repeat-pmatrix-label-layer");

        labelLayer.selectAll("text")
          .data(links.data())
          .enter()
          .append("text")
          .attr("x", function(d, i) {
            var key = d.source.name + "|" + d.target.name;
            var meta = lookup[key] || { pos: 0.5, dy: 0 };
            var path = links.nodes()[i];
            var point = path.getPointAtLength(path.getTotalLength() * meta.pos);
            return point.x;
          })
          .attr("y", function(d, i) {
            var key = d.source.name + "|" + d.target.name;
            var meta = lookup[key] || { pos: 0.5, dy: 0 };
            var path = links.nodes()[i];
            var point = path.getPointAtLength(path.getTotalLength() * meta.pos);
            return point.y + meta.dy;
          })
          .attr("text-anchor", "middle")
          .attr("dy", "0.1em")
          .style("font-size", "11px")
          .style("font-weight", "600")
          .style("fill", "#1f2937")
          .style("paint-order", "stroke")
          .style("stroke", "white")
          .style("stroke-width", "3px")
          .style("stroke-linejoin", "round")
          .style("pointer-events", "none")
          .text(function(d) {
            var key = d.source.name + "|" + d.target.name;
            return (lookup[key] || {}).band_label || "";
          });
      }, 300);
    }
    ',
    data = list(
      keys = paste(
        nodes_df$name[match(links_df$from, nodes_df$id)],
        nodes_df$name[match(links_df$to, nodes_df$id)],
        sep = "|"
      ),
      band_labels = links_df[[band_label_col]],
      observed_pcts = links_df$observed_pct_label,
      model_pcts = links_df$model_pct_label,
      counts = scales::comma(links_df$n_transitions),
      pos = links_df$label_pos,
      dy = links_df$label_dy,
      tooltips = links_df$tooltip_text
    )
  )
}

if (knitr::is_html_output() && requireNamespace("networkD3", quietly = TRUE)) {
  repeat_probability_sankey <- repeat_probability_sankey %>%
    dplyr::mutate(
      probability_value = 100 * probability,
      probability_tooltip = paste0(
        repeat_probability_sankey_nodes$name[match(from, repeat_probability_sankey_nodes$id)],
        " -> ",
        repeat_probability_sankey_nodes$name[match(to, repeat_probability_sankey_nodes$id)],
        "\nObserved donor share: ", observed_pct_label,
        "\nObserved transitions: ", scales::comma(n_transitions),
        "\nModel one-year probability: ", model_pct_label
      ),
      tooltip_text = probability_tooltip
    )

  repeat_observed_count_sankey <- repeat_probability_sankey %>%
    dplyr::mutate(
      to = dplyr::case_when(
        to_state == "Normoglycemia" ~ "observed_normoglycemia",
        to_state == "Prediabetes" ~ "observed_prediabetes",
        to_state == "Diabetes" ~ "observed_diabetes"
      ),
      count_label = scales::comma(n_transitions),
      tooltip_text = paste0(
        repeat_observed_count_sankey_nodes$name[match(from, repeat_observed_count_sankey_nodes$id)],
        " -> ",
        repeat_observed_count_sankey_nodes$name[match(to, repeat_observed_count_sankey_nodes$id)],
        "\nObserved transitions: ", scales::comma(n_transitions),
        "\nObserved donor share: ", observed_pct_label,
        "\nModel one-year probability: ", model_pct_label
      )
    )

  repeat_probability_sankey_widget <- build_repeat_sankey_widget(
    links_df = repeat_probability_sankey,
    nodes_df = repeat_probability_sankey_nodes,
    value_col = "probability_value",
    band_label_col = "observed_pct_label",
    units = "probability points"
  )

  repeat_observed_count_sankey_widget <- build_repeat_sankey_widget(
    links_df = repeat_observed_count_sankey,
    nodes_df = repeat_observed_count_sankey_nodes,
    value_col = "n_transitions",
    band_label_col = "count_label",
    units = "donors"
  )

  htmltools::tagList(
    htmltools::tags$div(
      style = "font-weight:600; margin: 0 0 8px 0;",
      "Probability-based Sankey with observed donor percentages"
    ),
    repeat_probability_sankey_widget,
    htmltools::tags$div(style = "height: 18px;"),
    htmltools::tags$div(
      style = "font-weight:600; margin: 0 0 8px 0;",
      "Observed-transition-count Sankey with donor counts"
    ),
    repeat_observed_count_sankey_widget
  )
} else {
  repeat_observed_count_sankey <- repeat_probability_sankey %>%
    dplyr::mutate(`Observed transitions` = scales::comma(n_transitions))

  htmltools::tagList(
    knitr::kable(
      repeat_probability_sankey %>%
        dplyr::transmute(
          `From state` = from_state,
          `To state` = to_state,
          `Model one-year probability` = model_pct_label,
          `Observed transitions` = scales::comma(n_transitions),
          `Observed share of starting state` = observed_pct_label
        ),
      caption = "Model-based one-year transition probabilities with observed transition shares from consecutive repeat-donor measurements. Render to HTML with networkD3 installed to view the interactive Sankey plot."
    ),
    knitr::kable(
      repeat_observed_count_sankey %>%
        dplyr::transmute(
          `From state` = from_state,
          `To observed state` = to_state,
          `Observed transitions` = `Observed transitions`,
          `Observed share of starting state` = observed_pct_label,
          `Model one-year probability` = model_pct_label
        ),
      caption = "Observed consecutive donor transitions between glycemic states. Render to HTML with networkD3 installed to view the interactive donor-count Sankey plot."
    )
  )
}


## ----repeat-a1c-race-effects, echo=FALSE, warning=FALSE, message=FALSE--------
if (inherits(repeat_race_msm, "error")) {
  cat(
    "Race-effect model unavailable in this session.\n\nReason:",
    conditionMessage(repeat_race_msm)
  )
} else {
  black_q <- msm::qmatrix.msm(
    repeat_race_msm,
    covariates = list(black_indicator = 1),
    ci = "none"
  )
  non_black_q <- msm::qmatrix.msm(
    repeat_race_msm,
    covariates = list(black_indicator = 0),
    ci = "none"
  )

  race_effect_tbl <- repeat_tidy_qmatrix(black_q) %>%
    dplyr::rename(Black = intensity) %>%
    dplyr::left_join(
      repeat_tidy_qmatrix(non_black_q) %>%
        dplyr::rename(`Non-Black` = intensity),
      by = c("from_id", "to_id", "transition")
    ) %>%
    dplyr::mutate(`Black/Non-Black rate ratio` = Black / `Non-Black`) %>%
    dplyr::arrange(from_id, to_id) %>%
    dplyr::select(transition, Black, `Non-Black`, `Black/Non-Black rate ratio`)

  gt::gt(race_effect_tbl) %>%
    gt::tab_header(
      title = "Effect of Race on Transition Intensities",
      subtitle = "Black (black_indicator = 1) versus non-Black (black_indicator = 0)"
    ) %>%
    gt::cols_label(
      transition = "Transition",
      Black = "Black",
      `Non-Black` = "Non-Black",
      `Black/Non-Black rate ratio` = "Black/Non-Black HR"
    ) %>%
    gt::tab_spanner(
      label = "Estimated intensity",
      columns = c(Black, `Non-Black`)
    ) %>%
    gt::fmt_number(columns = c(Black, `Non-Black`), decimals = 3) %>%
    gt::fmt_number(columns = `Black/Non-Black rate ratio`, decimals = 2) %>%
    gt::cols_align(
      align = "center",
      columns = c(Black, `Non-Black`, `Black/Non-Black rate ratio`)
    ) %>%
    gt::opt_row_striping() %>%
    gt::tab_source_note(
      source_note = "Off-diagonal entries from qmatrix.msm(); larger values indicate faster transitions between states."
    )
}


## ----repeat-a1c-urban-rural-effects-prep, echo=FALSE, warning=FALSE, message=FALSE, include=FALSE----
if (is.null(hotspot_map_data$error)) {
  repeat_urban_rural_model_data <- a1c_repeat_model_data %>%
    dplyr::left_join(hotspot_map_data$zcta_to_county, by = "zip5") %>%
    dplyr::left_join(hotspot_map_data$urban_rural_county, by = "GEOID") %>%
    dplyr::filter(urban_rural_group %in% c("Urban", "Rural")) %>%
    dplyr::mutate(
      urban_indicator = dplyr::case_when(
        urban_rural_group == "Urban" ~ 1,
        urban_rural_group == "Rural" ~ 0,
        TRUE ~ NA_real_
      )
    )

  repeat_urban_rural_msm <- tryCatch(
    msm::msm(
      status ~ cum_years,
      subject = id,
      data = repeat_urban_rural_model_data,
      qmatrix = repeat_q_init,
      deathexact = FALSE,
      exacttimes = TRUE,
      covariates = ~ urban_indicator,
      constraint = list(urban_indicator = c(1, 2, 3, 1, 2, 3)),
      control = list(trace = 0, REPORT = 1, maxit = 500)
    ),
    error = function(e) e
  )
}


## ----repeat-a1c-urban-rural-effects, echo=FALSE, warning=FALSE, message=FALSE----
if (!is.null(hotspot_map_data$error)) {
  cat(
    "Urban-rural-effect model unavailable in this session.\n\nReason:",
    hotspot_map_data$error
  )
} else if (inherits(repeat_urban_rural_msm, "error")) {
  cat(
    "Urban-rural-effect model unavailable in this session.\n\nReason:",
    conditionMessage(repeat_urban_rural_msm)
  )
} else {
  urban_q <- msm::qmatrix.msm(
    repeat_urban_rural_msm,
    covariates = list(urban_indicator = 1),
    ci = "none"
  )
  rural_q <- msm::qmatrix.msm(
    repeat_urban_rural_msm,
    covariates = list(urban_indicator = 0),
    ci = "none"
  )

  urban_rural_effect_tbl <- repeat_tidy_qmatrix(urban_q) %>%
    dplyr::rename(Urban = intensity) %>%
    dplyr::left_join(
      repeat_tidy_qmatrix(rural_q) %>%
        dplyr::rename(Rural = intensity),
      by = c("from_id", "to_id", "transition")
    ) %>%
    dplyr::mutate(`Rural/Urban rate ratio` = Rural / Urban) %>%
    dplyr::arrange(from_id, to_id) %>%
    dplyr::select(transition, Urban, Rural, `Rural/Urban rate ratio`)

  gt::gt(urban_rural_effect_tbl) %>%
    gt::tab_header(
      title = "Effect of Urbanicity on Transition Intensities",
      subtitle = "Urban (urban_indicator = 1) versus rural (urban_indicator = 0)"
    ) %>%
    gt::cols_label(
      transition = "Transition",
      Urban = "Urban",
      Rural = "Rural",
      `Rural/Urban rate ratio` = "Rural/Urban HR"
    ) %>%
    gt::tab_spanner(
      label = "Estimated intensity",
      columns = c(Urban, Rural)
    ) %>%
    gt::fmt_number(columns = c(Urban, Rural), decimals = 3) %>%
    gt::fmt_number(columns = `Rural/Urban rate ratio`, decimals = 2) %>%
    gt::cols_align(
      align = "center",
      columns = c(Urban, Rural, `Rural/Urban rate ratio`)
    ) %>%
    gt::opt_row_striping() %>%
    gt::tab_source_note(
      source_note = "Off-diagonal entries from qmatrix.msm(); larger values indicate faster transitions between states."
    )
}


## ----repeat-a1c-gender-effects, echo=FALSE, warning=FALSE, message=FALSE------
if (inherits(repeat_gender_msm, "error")) {
  cat(
    "Gender-effect model unavailable in this session.\n\nReason:",
    conditionMessage(repeat_gender_msm)
  )
} else {
  female_q <- msm::qmatrix.msm(
    repeat_gender_msm,
    covariates = list(gender_male = 0),
    ci = "none"
  )
  male_q <- msm::qmatrix.msm(
    repeat_gender_msm,
    covariates = list(gender_male = 1),
    ci = "none"
  )

  gender_effect_tbl <- repeat_tidy_qmatrix(female_q) %>%
    dplyr::rename(Female = intensity) %>%
    dplyr::left_join(
      repeat_tidy_qmatrix(male_q) %>%
        dplyr::rename(Male = intensity),
      by = c("from_id", "to_id", "transition")
    ) %>%
    dplyr::mutate(`Male/Female rate ratio` = Male / Female) %>%
    dplyr::arrange(from_id, to_id) %>%
    dplyr::select(transition, Female, Male, `Male/Female rate ratio`)

  gt::gt(gender_effect_tbl) %>%
    gt::tab_header(
      title = "Effect of Gender on Transition Intensities",
      subtitle = "Female (gender_male = 0) versus male (gender_male = 1)"
    ) %>%
    gt::cols_label(
      transition = "Transition",
      Female = "Female",
      Male = "Male",
      `Male/Female rate ratio` = "Male/Female HR"
    ) %>%
    gt::tab_spanner(
      label = "Estimated intensity",
      columns = c(Female, Male)
    ) %>%
    gt::fmt_number(columns = c(Female, Male), decimals = 3) %>%
    gt::fmt_number(columns = `Male/Female rate ratio`, decimals = 2) %>%
    gt::cols_align(
      align = "center",
      columns = c(Female, Male, `Male/Female rate ratio`)
    ) %>%
    gt::opt_row_striping() %>%
    gt::tab_source_note(
      source_note = "Off-diagonal entries from qmatrix.msm(); larger values indicate faster transitions between states."
    )
}

