args <- commandArgs(trailingOnly = TRUE)

default_input <- file.path("arc_a1c", "survey_a1c_demo.Rdata")
input_path <- if (length(args) >= 1) args[[1]] else default_input
output_dir <- if (length(args) >= 2) args[[2]] else "arc_a1c"

if (!file.exists(input_path)) {
  stop("Input file not found: ", normalizePath(input_path, winslash = "/", mustWork = FALSE))
}

dir.create(output_dir, recursive = TRUE, showWarnings = FALSE)

loaded_env <- new.env(parent = emptyenv())
loaded_names <- load(input_path, envir = loaded_env)

if (!"survey_a1c_demo" %in% loaded_names) {
  stop("Object 'survey_a1c_demo' was not found in ", basename(input_path))
}

survey_df <- get("survey_a1c_demo", envir = loaded_env)

question_vars <- c(
  "q1_a1c_knowledge",
  "q2_when_learned_a1c_testing",
  "q3_prior_a1c_awareness",
  "q4_interest_learning_a1c",
  "q4_scheduled_for_a1c_test",
  "q4_a1c_factored_donation_decision",
  "q4_a1c_motivates_future_donation",
  "q4_health_screenings_influence_donation"
)

missing_vars <- setdiff(c(question_vars, "glycemic_category"), names(survey_df))
if (length(missing_vars) > 0) {
  stop("Required variables missing from survey data: ", paste(missing_vars, collapse = ", "))
}

question_order <- list(
  q1_a1c_knowledge = c("Yes", "No", "Unsure"),
  q2_when_learned_a1c_testing = c("Aware Prior to Donation", "Learned at Donation", "Unaware"),
  q3_prior_a1c_awareness = c("Yes", "No", "Unsure"),
  q4_interest_learning_a1c = c("Strongly Agree", "Agree", "Neutral", "Disagree", "Strongly Disagree"),
  q4_scheduled_for_a1c_test = c("Strongly Agree", "Agree", "Neutral", "Disagree", "Strongly Disagree"),
  q4_a1c_factored_donation_decision = c("Strongly Agree", "Agree", "Neutral", "Disagree", "Strongly Disagree"),
  q4_a1c_motivates_future_donation = c("Strongly Agree", "Agree", "Neutral", "Disagree", "Strongly Disagree"),
  q4_health_screenings_influence_donation = c("Strongly Agree", "Agree", "Neutral", "Disagree", "Strongly Disagree")
)

question_labels <- vapply(
  question_vars,
  function(var_name) {
    label_value <- attr(survey_df[[var_name]], "label")
    if (is.null(label_value) || !nzchar(label_value)) {
      var_name
    } else {
      label_value
    }
  },
  character(1)
)

question_results <- lapply(question_vars, function(var_name) {
  response_values <- as.character(survey_df[[var_name]])
  glycemic_values <- as.character(survey_df$glycemic_category)

  valid_rows <- !is.na(response_values) &
    nzchar(response_values) &
    !is.na(glycemic_values) &
    nzchar(glycemic_values)

  response_values <- response_values[valid_rows]
  glycemic_values <- glycemic_values[valid_rows]

  response_levels <- question_order[[var_name]]
  observed_responses <- unique(response_values)
  response_levels <- unique(c(response_levels, sort(setdiff(observed_responses, response_levels))))

  response_factor <- factor(response_values, levels = response_levels)
  glycemic_factor <- factor(
    glycemic_values,
    levels = c("Normoglycemia", "Prediabetes", "Diabetes")
  )

  counts_matrix <- table(response_factor, glycemic_factor)
  counts_df <- as.data.frame.matrix(counts_matrix, stringsAsFactors = FALSE)

  for (gly_col in c("Normoglycemia", "Prediabetes", "Diabetes")) {
    if (!gly_col %in% names(counts_df)) {
      counts_df[[gly_col]] <- 0L
    }
  }

  counts_df$response <- rownames(counts_df)
  rownames(counts_df) <- NULL

  counts_df$response_n <- rowSums(counts_df[, c("Normoglycemia", "Prediabetes", "Diabetes"), drop = FALSE])
  counts_df <- counts_df[counts_df$response_n > 0, , drop = FALSE]

  counts_df$normoglycemia_pct <- round(100 * counts_df$Normoglycemia / counts_df$response_n, 1)
  counts_df$prediabetes_pct <- round(100 * counts_df$Prediabetes / counts_df$response_n, 1)
  counts_df$diabetes_pct <- round(100 * counts_df$Diabetes / counts_df$response_n, 1)

  counts_df$question_var <- var_name
  counts_df$question_label <- question_labels[[var_name]]
  counts_df$question_n <- sum(counts_df$response_n)

  chisq_p <- tryCatch(
    suppressWarnings(chisq.test(counts_matrix[counts_df$response, , drop = FALSE])$p.value),
    error = function(e) NA_real_
  )
  counts_df$chisq_p_value <- chisq_p

  counts_df[, c(
    "question_var",
    "question_label",
    "question_n",
    "response",
    "response_n",
    "Normoglycemia",
    "Prediabetes",
    "Diabetes",
    "normoglycemia_pct",
    "prediabetes_pct",
    "diabetes_pct",
    "chisq_p_value"
  )]
})

distribution_long <- do.call(rbind, question_results)

distribution_wide <- distribution_long
names(distribution_wide)[names(distribution_wide) == "Normoglycemia"] <- "normoglycemia_n"
names(distribution_wide)[names(distribution_wide) == "Prediabetes"] <- "prediabetes_n"
names(distribution_wide)[names(distribution_wide) == "Diabetes"] <- "diabetes_n"

summary_by_question <- do.call(
  rbind,
  lapply(split(distribution_long, distribution_long$question_var), function(df_question) {
    highest_diabetes <- df_question[which.max(df_question$diabetes_pct), , drop = FALSE]
    lowest_diabetes <- df_question[which.min(df_question$diabetes_pct), , drop = FALSE]

    data.frame(
      question_var = df_question$question_var[[1]],
      question_label = df_question$question_label[[1]],
      question_n = df_question$question_n[[1]],
      chisq_p_value = df_question$chisq_p_value[[1]],
      highest_diabetes_response = highest_diabetes$response[[1]],
      highest_diabetes_pct = highest_diabetes$diabetes_pct[[1]],
      lowest_diabetes_response = lowest_diabetes$response[[1]],
      lowest_diabetes_pct = lowest_diabetes$diabetes_pct[[1]],
      stringsAsFactors = FALSE
    )
  })
)

long_path <- file.path(output_dir, "survey_question_glycemic_distribution_long.csv")
wide_path <- file.path(output_dir, "survey_question_glycemic_distribution_wide.csv")
summary_path <- file.path(output_dir, "survey_question_glycemic_distribution_summary.csv")

utils::write.csv(distribution_long, long_path, row.names = FALSE, na = "")
utils::write.csv(distribution_wide, wide_path, row.names = FALSE, na = "")
utils::write.csv(summary_by_question, summary_path, row.names = FALSE, na = "")

if (requireNamespace("openxlsx", quietly = TRUE)) {
  workbook_path <- file.path(output_dir, "survey_question_glycemic_distribution.xlsx")
  wb <- openxlsx::createWorkbook()

  openxlsx::addWorksheet(wb, "All questions")
  openxlsx::writeData(wb, "All questions", distribution_wide)

  openxlsx::addWorksheet(wb, "Question summary")
  openxlsx::writeData(wb, "Question summary", summary_by_question)

  for (var_name in question_vars) {
    sheet_name <- gsub("[\\\\/:*?\\[\\]]", " ", substr(var_name, 1, 31))
    openxlsx::addWorksheet(wb, sheet_name)
    openxlsx::writeData(
      wb,
      sheet_name,
      distribution_wide[distribution_wide$question_var == var_name, , drop = FALSE]
    )
  }

  openxlsx::saveWorkbook(wb, workbook_path, overwrite = TRUE)
  cat("Wrote:", normalizePath(workbook_path, winslash = "/", mustWork = TRUE), "\n")
}

cat("Wrote:", normalizePath(long_path, winslash = "/", mustWork = TRUE), "\n")
cat("Wrote:", normalizePath(wide_path, winslash = "/", mustWork = TRUE), "\n")
cat("Wrote:", normalizePath(summary_path, winslash = "/", mustWork = TRUE), "\n")
