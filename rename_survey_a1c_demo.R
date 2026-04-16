args <- commandArgs(trailingOnly = TRUE)

default_path <- file.path("arc_a1c", "survey_a1c_demo.Rdata")
target_path <- if (length(args) >= 1) args[[1]] else default_path

if (!file.exists(target_path)) {
  stop("Rdata file not found: ", normalizePath(target_path, winslash = "/", mustWork = FALSE))
}

loaded_env <- new.env(parent = emptyenv())
loaded_names <- load(target_path, envir = loaded_env)

if ("survey_a1c_demo" %in% loaded_names) {
  object_name <- "survey_a1c_demo"
} else {
  data_objects <- loaded_names[vapply(
    loaded_names,
    function(x) inherits(get(x, envir = loaded_env), c("data.frame", "tbl_df")),
    logical(1)
  )]

  if (length(data_objects) != 1) {
    stop(
      "Expected a single survey data.frame in the Rdata file or an object named 'survey_a1c_demo'. ",
      "Found: ", paste(loaded_names, collapse = ", ")
    )
  }

  object_name <- data_objects[[1]]
}

survey_df <- get(object_name, envir = loaded_env)

old_to_new <- c(
  "X1..Do.you.know.what.Hemoglobin.A1C.testing.is." =
    "q1_a1c_knowledge",
  "X2..Which.of.the.following.statements.best.describes.when.you.learned.that.the.Red.Cross.was.offering.Hemoglobin.A1C.testing.as.part.of.the.blood.donation.process.in.March." =
    "q2_when_learned_a1c_testing",
  "X3..Prior.to.your.recent.donation.were.you.aware.of.your.A1C.level." =
    "q3_prior_a1c_awareness",
  "X4.1.Learning.my.A1C.level.through.a.blood.donation.is.of.interest.to.me." =
    "q4_interest_learning_a1c",
  "X4.2.I.scheduled.an.appointment.to.give.so.that.I.could.receive.a.complimentary.A1C.test." =
    "q4_scheduled_for_a1c_test",
  "X4.3.Knowing.about.complimentary.A1C.testing.was.a.factor.in.my.decision.to.give.blood.or.platelets." =
    "q4_a1c_factored_donation_decision",
  "X4.4.Hemoglobin.A1C.testing.would.motivate.me.to.schedule.a.donation.appointment.in.the.future." =
    "q4_a1c_motivates_future_donation",
  "X4.5.Receiving.complementary.health.screenings..e.g..A1C.blood.pressure.etc...will.influence.me.to.donate.blood.in.the.future." =
    "q4_health_screenings_influence_donation",
  "X4.1.Learning.my.A1C.level.through.a.blood.donation.is.of.interest.to.me..1" =
    "q4_interest_learning_a1c_score",
  "X4.2.I.scheduled.an.appointment.to.give.so.that.I.could.receive.a.complimentary.A1C.test..1" =
    "q4_scheduled_for_a1c_test_score",
  "X4.3.Knowing.about.complimentary.A1C.testing.was.a.factor.in.my.decision.to.give.blood.or.platelets.during.the.month.of.March." =
    "q4_a1c_factored_donation_decision_score",
  "X4.4.Hemoglobin.A1C.testing.would.motivate.me.to.schedule.a.donation.appointment.in.the.future..1" =
    "q4_a1c_motivates_future_donation_score",
  "X4.5.Receiving.complementary.health.screenings.will.influence.me.to.donate.blood.in.the.future." =
    "q4_health_screenings_influence_donation_score"
)

label_map <- c(
  q1_a1c_knowledge = "Do you know what Hemoglobin A1C testing is?",
  q2_when_learned_a1c_testing = paste(
    "Which of the following statements best describes when you learned",
    "that the Red Cross was offering Hemoglobin A1C testing as part of",
    "the blood donation process in March?"
  ),
  q3_prior_a1c_awareness = "Prior to your recent donation were you aware of your A1C level?",
  q4_interest_learning_a1c = "Learning my A1C level through a blood donation is of interest to me.",
  q4_scheduled_for_a1c_test = paste(
    "I scheduled an appointment to give so that I could receive a",
    "complimentary A1C test."
  ),
  q4_a1c_factored_donation_decision = paste(
    "Knowing about complimentary A1C testing was a factor in my decision",
    "to give blood or platelets."
  ),
  q4_a1c_motivates_future_donation = paste(
    "Hemoglobin A1C testing would motivate me to schedule a donation",
    "appointment in the future."
  ),
  q4_health_screenings_influence_donation = paste(
    "Receiving complementary health screenings (e.g., A1C, blood pressure, etc.)",
    "will influence me to donate blood in the future."
  ),
  q4_interest_learning_a1c_score = "Interest in learning A1C through donation score",
  q4_scheduled_for_a1c_test_score = "Scheduled appointment for complimentary A1C test score",
  q4_a1c_factored_donation_decision_score = "Complimentary A1C testing factored into donation decision score",
  q4_a1c_motivates_future_donation_score = "A1C testing motivates future donation score",
  q4_health_screenings_influence_donation_score = "Complementary health screenings influence future donation score"
)

rename_count <- 0L
already_present <- character(0)
conflicts <- character(0)

for (old_name in names(old_to_new)) {
  new_name <- old_to_new[[old_name]]

  if (old_name %in% names(survey_df) && !(new_name %in% names(survey_df))) {
    names(survey_df)[names(survey_df) == old_name] <- new_name
    rename_count <- rename_count + 1L
  } else if (new_name %in% names(survey_df) && !(old_name %in% names(survey_df))) {
    already_present <- c(already_present, new_name)
  } else if (old_name %in% names(survey_df) && new_name %in% names(survey_df)) {
    conflicts <- c(conflicts, sprintf("%s -> %s", old_name, new_name))
  }
}

for (var_name in names(label_map)) {
  if (var_name %in% names(survey_df)) {
    attr(survey_df[[var_name]], "label") <- label_map[[var_name]]
  }
}

attr(survey_df, "variable_label_map") <- label_map
attr(survey_df, "rename_map_from_raw") <- old_to_new

assign(object_name, survey_df, envir = loaded_env)
save(list = loaded_names, file = target_path, envir = loaded_env)

cat("Updated object:", object_name, "\n")
cat("File:", normalizePath(target_path, winslash = "/", mustWork = TRUE), "\n")
cat("Columns renamed in this run:", rename_count, "\n")
cat("Columns already in final form:", length(already_present), "\n")

if (length(conflicts) > 0) {
  warning(
    "Some raw and renamed columns were both present and were left unchanged: ",
    paste(conflicts, collapse = "; ")
  )
}
