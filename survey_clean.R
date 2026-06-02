# survey_clean.R
#
# Recodes free-text and county responses from the survey frequencies output.
# Produces a cleaned version of the county_freq and free_text sheets.
#
# Install packages once:
#   install.packages(c("readxl", "writexl", "dplyr", "stringr"))
#
# Workflow:
#   1. Source this file once to see the raw values
#   2. Fill in the recode rules in the sections below
#   3. Source again — cleaned output is written to survey_frequencies_cleaned.xlsx

library(readxl)
library(writexl)
library(dplyr)
library(stringr)

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE <- file.path(path.expand("~"), "Downloads", "CRC MDH Project")

FREQ_FILE    <- file.path(BASE, "MCHHealthcareProvide-DataSetForLauraAndNo_DATA_LABELS_2026-06-01_1832_frequencies.xlsx")
OUTPUT_FILE  <- sub("\\.xlsx$", "_cleaned.xlsx", FREQ_FILE)

# ── Read ──────────────────────────────────────────────────────────────────────
county_raw  <- read_excel(FREQ_FILE, sheet = "county")
free_text   <- read_excel(FREQ_FILE, sheet = "free_text")

# ── County recoding ───────────────────────────────────────────────────────────
# Run once to see all unique raw values:
cat("\n── Raw county values ──────────────────────────────────────────\n")
county_raw %>%
  count(county_raw, sort = TRUE) %>%
  print(n = Inf)

# Add recode rules here using case_when.
# Left side  = pattern to match (case-insensitive substring)
# Right side = standardized name to use
county_recoded <- county_raw %>%
  mutate(
    county_clean = case_when(
      # ── Add your rules below ────────────────────────────────────────────────
      # str_detect(str_to_lower(county_raw), "hennepin")          ~ "Hennepin",
      # str_detect(str_to_lower(county_raw), "ramsey")            ~ "Ramsey",
      # str_detect(str_to_lower(county_raw), "st\\.?\\s*louis")   ~ "St Louis",
      # str_detect(str_to_lower(county_raw), "dakota")            ~ "Dakota",
      # ── Fallback: use the normalized value from Python ──────────────────────
      TRUE ~ county_normalized
    )
  )

# Summary of cleaned county counts by eligibility
county_summary <- county_recoded %>%
  group_by(county_clean, group) %>%
  summarise(n = n(), .groups = "drop") %>%
  tidyr::pivot_wider(names_from = group, values_from = n, values_fill = 0) %>%
  mutate(total = rowSums(across(where(is.numeric)))) %>%
  arrange(desc(total))

cat("\n── County summary (after recoding) ────────────────────────────\n")
print(county_summary, n = Inf)

# ── Free-text recoding ────────────────────────────────────────────────────────
# See all unique questions with free-text responses
cat("\n── Free-text questions ─────────────────────────────────────────\n")
free_text %>%
  count(question, sort = TRUE) %>%
  print(n = Inf)

# To recode a specific question, filter and apply case_when.
# Example for a "which other" field:
#
# free_text_recoded <- free_text %>%
#   mutate(
#     response_clean = case_when(
#       str_detect(question, "other.*screen") ~ case_when(
#         str_detect(str_to_lower(response), "always")  ~ "Always screens",
#         str_detect(str_to_lower(response), "never")   ~ "Never screens",
#         TRUE ~ response   # keep as-is if no rule matches
#       ),
#       TRUE ~ response
#     )
#   )

# For now, pass through unchanged — add rules above then re-source
free_text_recoded <- free_text %>% mutate(response_clean = response)

# ── Write cleaned output ──────────────────────────────────────────────────────
sheets_out <- list(
  county          = county_recoded,
  county_summary  = county_summary,
  free_text       = free_text_recoded
)

write_xlsx(sheets_out, OUTPUT_FILE)
cat("\nCleaned file saved to:", basename(OUTPUT_FILE), "\n")
