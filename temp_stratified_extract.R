setwd("e:/ARC/Projects/A1C/data/all/arc_a1c")

suppressPackageStartupMessages({
  library(dplyr)
})

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
  )

dedup <- a1c_all_a1cdemo %>%
  arrange(.data[["id"]], desc(.data[["Donation Date"]])) %>%
  distinct(.data[["id"]], .keep_all = TRUE)

df1 <- dedup %>%
  filter(!is.na(gender)) %>%
  group_by(age_group, gender, glycemic_category) %>%
  count()

df2 <- dedup %>%
  filter(!is.na(gender)) %>%
  group_by(age_group, gender) %>%
  count()

out <- df1 %>%
  left_join(df2, by = c("age_group", "gender")) %>%
  mutate(percentage = round(100 * n.x / n.y, 1)) %>%
  select(age_group, gender, glycemic_category, percentage) %>%
  arrange(age_group, gender, glycemic_category)

print(out, n = Inf)
