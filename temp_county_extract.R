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

load("a1c_all_a1cdemo.Rdata")

a1c_all_a1cdemo <- a1c_all_a1cdemo %>%
  mutate(
    `Donation Date` = as.Date(.data[["Donation Date"]]),
    `ARC A1C Results` = as.numeric(.data[["ARC A1C Results"]]),
    glycemic_category = case_when(
      is.na(.data[["ARC A1C Results"]]) ~ NA_character_,
      .data[["ARC A1C Results"]] < 5.7 ~ "Normoglycemia",
      .data[["ARC A1C Results"]] >= 5.7 & .data[["ARC A1C Results"]] < 6.5 ~ "Prediabetes",
      .data[["ARC A1C Results"]] >= 6.5 ~ "Diabetes"
    )
  ) %>%
  rename(a1c = `ARC A1C Results`)

dedup <- a1c_all_a1cdemo %>%
  arrange(.data[["id"]], desc(.data[["Donation Date"]])) %>%
  distinct(.data[["id"]], .keep_all = TRUE)

tryCatch({
  diabetes_geo <- dedup %>%
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

  cat("zip_prefixes=", length(unique(substr(diabetes_geo$zip5, 1, 3))), "\n", sep = "")

  zcta_sf <- tigris::zctas(
    cb = TRUE,
    starts_with = sort(unique(substr(diabetes_geo$zip5, 1, 3))),
    year = 2020,
    class = "sf"
  )
  cat("zcta_n=", nrow(zcta_sf), "\n", sep = "")

  counties_sf <- tigris::counties(cb = TRUE, year = 2020, class = "sf") %>%
    filter(!STATEFP %in% c("60", "66", "69", "72", "78"))
  states_sf <- tigris::states(cb = TRUE, year = 2020, class = "sf") %>%
    filter(!STUSPS %in% c("AS", "GU", "MP", "PR", "VI"))
  state_lookup <- states_sf %>% st_drop_geometry() %>% select(STATEFP, STUSPS)
  cat("counties_n=", nrow(counties_sf), "\n", sep = "")

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
  cat("zip_id_col=", zip_id_col, "\n", sep = "")

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
  cat("zcta_to_county_n=", nrow(zcta_to_county), "\n", sep = "")

  county_diabetes <- diabetes_geo %>%
    left_join(zcta_to_county, by = "zip5") %>%
    filter(!is.na(GEOID)) %>%
    group_by(GEOID) %>%
    summarise(
      n_donors = n(),
      diabetes_prev = 100 * mean(diabetes_flag, na.rm = TRUE),
      .groups = "drop"
    )
  cat("county_diabetes_n=", nrow(county_diabetes), "\n", sep = "")

  hotspot_input <- counties_sf %>%
    left_join(county_diabetes, by = "GEOID") %>%
    filter(!is.na(diabetes_prev), n_donors >= 30)
  cat("hotspot_input_n=", nrow(hotspot_input), "\n", sep = "")

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

  svi_test <- cor.test(
    county_assoc$diabetes_prev,
    county_assoc$svi_overall,
    method = "spearman",
    exact = FALSE
  )
  cat("spearman_rho=", round(unname(svi_test$estimate), 3), "\n", sep = "")
  cat("spearman_p=", format.pval(svi_test$p.value), "\n", sep = "")
  cat("diabetes_prev_range=", paste(round(range(county_assoc$diabetes_prev, na.rm = TRUE), 2), collapse = " to "), "\n", sep = "")
}, error = function(e) {
  cat("ERROR:", conditionMessage(e), "\n")
  quit(status = 1)
})
