setwd("e:/ARC/Projects/A1C/data/all/arc_a1c")

suppressPackageStartupMessages({
  library(dplyr)
  library(stringr)
  library(readr)
  library(readxl)
  library(sf)
  library(spdep)
  library(tigris)
})

options(tigris_use_cache = TRUE)

load("a1c_all.Rdata")
load("a1c_all_a1cdemo.Rdata")

medical_history_cols <- c(
  "cancer", "cardiorespiratory", "coagulopathy", "teratogen",
  "anticoagulants", "deferral_meds", "hypertension", "unwell",
  "splenectomy_tcpenia", "read_edu_mat"
)

a1c_all_a1cdemo <- a1c_all_a1cdemo %>%
  mutate(
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
  mutate(
    across(all_of(medical_history_cols), ~ ifelse(is.na(.x), 0, .x)),
    `ARC A1C Results` = as.numeric(.data[["ARC A1C Results"]]),
    glycemic_category = case_when(
      is.na(.data[["ARC A1C Results"]]) ~ NA_character_,
      .data[["ARC A1C Results"]] < 5.7 ~ "Normoglycemia",
      .data[["ARC A1C Results"]] >= 5.7 & .data[["ARC A1C Results"]] < 6.5 ~ "Prediabetes",
      .data[["ARC A1C Results"]] >= 6.5 ~ "Diabetes"
    ),
    glycemic_category = factor(
      glycemic_category,
      levels = c("Normoglycemia", "Prediabetes", "Diabetes")
    ),
    period = case_when(
      .data[["Donation Date"]] >= as.Date("2025-03-01") & .data[["Donation Date"]] <= as.Date("2025-04-30") ~ "Mar 2025",
      .data[["Donation Date"]] >= as.Date("2025-07-01") & .data[["Donation Date"]] <= as.Date("2025-09-30") ~ "Aug 2025",
      .data[["Donation Date"]] >= as.Date("2025-11-01") & .data[["Donation Date"]] <= as.Date("2025-12-31") ~ "Nov 2025",
      .data[["Donation Date"]] >= as.Date("2026-03-01") & .data[["Donation Date"]] <= as.Date("2026-04-30") ~ "Mar 2026",
      TRUE ~ NA_character_
    )
  ) %>%
  rename(a1c = `ARC A1C Results`) %>%
  mutate(
    raceeth = case_when(
      race %in% c("Caucasian") ~ "White",
      race %in% c("African American") ~ "Black",
      race %in% c("Hispanic") ~ "Hispanic Origin",
      race %in% c("Asian", "Mix", "Native American", "Other", "Prefer not to answer", "") ~ "Asian/AIAN/other",
      TRUE ~ ""
    )
  )

a1c_all_a1cdemo_dedup <- a1c_all_a1cdemo %>%
  arrange(.data[["id"]], desc(.data[["Donation Date"]])) %>%
  distinct(.data[["id"]], .keep_all = TRUE)

a1c_all <- a1c_all %>%
  mutate(
    `Donation Date` = as.Date(.data[["Donation Date"]]),
    `ARC A1C Results` = as.numeric(.data[["ARC A1C Results"]]),
    glycemic_category = case_when(
      is.na(.data[["ARC A1C Results"]]) ~ NA_character_,
      .data[["ARC A1C Results"]] < 5.7 ~ "Normoglycemia",
      .data[["ARC A1C Results"]] >= 5.7 & .data[["ARC A1C Results"]] < 6.5 ~ "Prediabetes",
      .data[["ARC A1C Results"]] >= 6.5 ~ "Diabetes"
    ),
    glycemic_category = factor(
      glycemic_category,
      levels = c("Normoglycemia", "Prediabetes", "Diabetes")
    ),
    period = case_when(
      .data[["Donation Date"]] >= as.Date("2025-03-01") & .data[["Donation Date"]] <= as.Date("2025-04-30") ~ "Mar 2025",
      .data[["Donation Date"]] >= as.Date("2025-07-01") & .data[["Donation Date"]] <= as.Date("2025-09-30") ~ "Aug 2025",
      .data[["Donation Date"]] >= as.Date("2025-11-01") & .data[["Donation Date"]] <= as.Date("2025-12-31") ~ "Nov 2025",
      .data[["Donation Date"]] >= as.Date("2026-03-01") & .data[["Donation Date"]] <= as.Date("2026-04-30") ~ "Mar 2026",
      TRUE ~ NA_character_
    )
  )

cat("GLYCEMIC BY PERIOD\n")
glycemic_period_counts <- a1c_all %>%
  filter(!is.na(period), !is.na(glycemic_category)) %>%
  arrange(.data[["period"]], .data[["id"]], .data[["Donation Date"]], .data[["Donation Id"]]) %>%
  distinct(period, id, .keep_all = TRUE) %>%
  count(period, glycemic_category, name = "distinct_donors") %>%
  group_by(period) %>%
  mutate(period_total = sum(distinct_donors), pct = distinct_donors / period_total) %>%
  ungroup()
print(glycemic_period_counts)

cat("\nAGE/GENDER STRATIFIED\n")
df_a1ccatcount <- a1c_all_a1cdemo_dedup %>%
  filter(!is.na(gender)) %>%
  group_by(age_group, gender, glycemic_category) %>%
  count()
df_count <- a1c_all_a1cdemo_dedup %>%
  filter(!is.na(gender)) %>%
  group_by(age_group, gender) %>%
  count()
df_diab_p <- df_a1ccatcount %>%
  left_join(df_count, by = c("age_group", "gender")) %>%
  mutate(percentage = round(n.x / n.y, digits = 3)) %>%
  rename(group_count = n.x, total_donors = n.y)
print(df_diab_p %>% select(age_group, gender, glycemic_category, group_count, total_donors, percentage))

glycemic_tab <- xtabs(
  ~ interaction(age_group, gender) + glycemic_category,
  data = a1c_all_a1cdemo_dedup %>%
    filter(!is.na(age_group), !is.na(gender), !is.na(glycemic_category), glycemic_category != "hypoglycemia") %>%
    mutate(glycemic_category = droplevels(factor(glycemic_category)))
)
glycemic_chisq <- chisq.test(glycemic_tab)
cat("chisq_p=", format.pval(glycemic_chisq$p.value), "\n", sep = "")

cat("\nCOUNTY SUMMARY\n")
tryCatch({
  diabetes_geo <- a1c_all_a1cdemo_dedup %>%
    filter(!is.na(glycemic_category)) %>%
    mutate(
      zip5 = str_pad(
        str_extract(as.character(zip), "\\d{1,5}"),
        width = 5,
        side = "left",
        pad = "0"
      ),
      diabetes_flag = as.integer(glycemic_category == "Diabetes"),
      age_55_plus = case_when(
        !is.na(age) ~ as.numeric(age >= 55),
        age_group == "55-" ~ 1,
        age_group %in% c("-29", "30-54") ~ 0,
        TRUE ~ NA_real_
      )
    ) %>%
    filter(!is.na(zip5))

  zip_prefixes <- sort(unique(substr(diabetes_geo$zip5, 1, 3)))
  zcta_sf <- tigris::zctas(cb = TRUE, starts_with = zip_prefixes, year = 2020, class = "sf")
  counties_sf <- tigris::counties(cb = TRUE, year = 2020, class = "sf") %>%
    filter(!STATEFP %in% c("60", "66", "69", "72", "78"))
  states_sf <- tigris::states(cb = TRUE, year = 2020, class = "sf") %>%
    filter(!STUSPS %in% c("AS", "GU", "MP", "PR", "VI"))
  state_lookup <- states_sf %>% st_drop_geometry() %>% select(STATEFP, STUSPS)

  format_county_geoid <- function(x) {
    x_chr <- trimws(as.character(x))
    x_chr <- sub("\\.0$", "", x_chr)
    str_pad(x_chr, width = 5, side = "left", pad = "0")
  }

  urban_rural_county <- readxl::read_excel("urban rural classification CDC.xlsx") %>%
    {
      names(.)[names(.) == "2023 Code"] <- "urban_code"
      .
    } %>%
    transmute(
      GEOID = format_county_geoid(Location),
      urban_rural_code = as.character(urban_code),
      urban_rural_group = case_when(
        substr(urban_rural_code, 1, 1) %in% c("1", "2", "3", "4") ~ "Urban",
        substr(urban_rural_code, 1, 1) %in% c("5", "6") ~ "Rural",
        TRUE ~ "Unknown"
      )
    )

  svi_county <- readr::read_csv("SVI_2022_US_county.csv", show_col_types = FALSE) %>%
    transmute(
      GEOID = format_county_geoid(FIPS),
      svi_overall = as.numeric(RPL_THEMES)
    )

  zip_id_col <- intersect(c("ZCTA5CE20", "GEOID20", "ZCTA5CE10"), names(zcta_sf))[1]

  zcta_to_county <- zcta_sf %>%
    st_transform(5070) %>%
    st_point_on_surface() %>%
    st_transform(st_crs(counties_sf)) %>%
    st_join(counties_sf %>% select(GEOID), left = FALSE) %>%
    st_drop_geometry() %>%
    transmute(
      zip5 = .data[[zip_id_col]],
      GEOID = .data[["GEOID"]]
    ) %>%
    distinct(zip5, .keep_all = TRUE)

  county_diabetes <- diabetes_geo %>%
    left_join(zcta_to_county, by = "zip5") %>%
    filter(!is.na(GEOID)) %>%
    group_by(GEOID) %>%
    summarise(
      n_donors = n(),
      diabetes_cases = sum(diabetes_flag, na.rm = TRUE),
      diabetes_prev = 100 * mean(diabetes_flag, na.rm = TRUE),
      median_age = if (all(is.na(age))) NA_real_ else median(age, na.rm = TRUE),
      pct_age_55_plus = if (all(is.na(age_55_plus))) NA_real_ else 100 * mean(age_55_plus, na.rm = TRUE),
      .groups = "drop"
    )

  hotspot_input <- counties_sf %>%
    left_join(county_diabetes, by = "GEOID") %>%
    filter(!is.na(diabetes_prev), n_donors >= 30)

  county_nb <- spdep::poly2nb(hotspot_input, queen = TRUE)
  county_lw <- spdep::nb2listw(county_nb, style = "W", zero.policy = TRUE)
  hotspot_input$gi_star <- as.numeric(
    spdep::localG(hotspot_input$diabetes_prev, county_lw, zero.policy = TRUE)
  )

  hotspot_class <- hotspot_input %>%
    st_drop_geometry() %>%
    transmute(
      GEOID,
      hotspot = case_when(
        gi_star >= 1.96 ~ "Hot spot",
        gi_star <= -1.96 ~ "Cold spot",
        TRUE ~ "Not significant"
      )
    )

  county_map <- counties_sf %>%
    {
      if ("STUSPS" %in% names(.)) . else left_join(., state_lookup, by = "STATEFP")
    } %>%
    left_join(county_diabetes, by = "GEOID") %>%
    left_join(urban_rural_county, by = "GEOID") %>%
    left_join(svi_county, by = "GEOID") %>%
    left_join(hotspot_class, by = "GEOID") %>%
    mutate(
      hotspot = case_when(
        is.na(n_donors) ~ "No donor data",
        n_donors < 30 ~ "Insufficient donor count",
        TRUE ~ hotspot
      )
    )

  county_assoc <- county_map %>%
    st_drop_geometry() %>%
    filter(!is.na(diabetes_prev), !is.na(svi_overall), n_donors >= 30)

  print(table(county_map$hotspot, useNA = "ifany"))
  cat("n_counties_ge30=", nrow(county_assoc), "\n", sep = "")
  svi_test <- cor.test(county_assoc$diabetes_prev, county_assoc$svi_overall, method = "spearman", exact = FALSE)
  cat("spearman_rho=", round(unname(svi_test$estimate), 3), "\n", sep = "")
  cat("spearman_p=", format.pval(svi_test$p.value), "\n", sep = "")
  cat("diabetes_prev_range=", paste(round(range(county_assoc$diabetes_prev, na.rm = TRUE), 2), collapse = " to "), "\n", sep = "")
}, error = function(e) {
  cat("COUNTY_ERROR=", conditionMessage(e), "\n", sep = "")
})
