# survey_report.R
#
# Reads survey_frequencies.xlsx and writes a formatted Word report.
# Each question gets its own heading and clean frequency table.
#
# Install packages once (paste into RStudio console):
#   install.packages(c("readxl", "officer", "flextable", "dplyr", "stringr"))
#
# Then open this file in RStudio and click Source (Ctrl+Shift+S / Cmd+Shift+S)

library(readxl)
library(officer)
library(flextable)
library(dplyr)
library(stringr)

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE <- file.path(path.expand("~"), "Downloads", "CRC MDH Project", "MDH analysis")

# Change this to your frequencies file name
FREQ_FILE   <- file.path(BASE, "MCHHealthcareProvide-DataSetForLauraAndNo_DATA_LABELS_2026-06-01_1832_frequencies.xlsx")
OUTPUT_FILE <- sub("\\.xlsx$", "_report.docx", FREQ_FILE)

cat("Input :", basename(FREQ_FILE), "\n")
cat("Output:", basename(OUTPUT_FILE), "\n")

# ── Read sheets ───────────────────────────────────────────────────────────────
available <- excel_sheets(FREQ_FILE)

read_if <- function(name) {
  if (name %in% available) read_excel(FREQ_FILE, sheet = name) else NULL
}

eligible    <- read_if("eligible")
ineligible  <- read_if("ineligible")
county_freq <- read_if("county_freq")

# ── Flextable helper ──────────────────────────────────────────────────────────
make_table <- function(df) {
  # Rename N columns for display
  df <- df %>% rename_with(~ "N", any_of(c("N (answered)", "N (denominator)")))

  ft <- flextable(df) %>%
    theme_booktabs() %>%
    bold(part = "header") %>%
    fontsize(size = 10, part = "all") %>%
    font(fontname = "Calibri", part = "all") %>%
    align(j = c("n", "%", "N"), align = "right", part = "all") %>%
    width(j = "Response", width = 3.5) %>%
    width(j = c("n", "%", "N"), width = 0.7)

  ft
}

make_county_table <- function(df) {
  flextable(df) %>%
    theme_booktabs() %>%
    bold(part = "header") %>%
    bold(i = nrow(df), part = "body") %>%   # bold TOTAL row
    fontsize(size = 10, part = "all") %>%
    font(fontname = "Calibri", part = "all") %>%
    align(j = -1, align = "right", part = "all") %>%
    autofit()
}

# ── Add a frequency section to the doc ───────────────────────────────────────
# df must have columns: Question, Response, n, %, and N (answered) or N (denominator)
add_freq_section <- function(doc, df, heading) {
  if (is.null(df) || nrow(df) == 0) return(doc)

  doc <- body_add_par(doc, heading, style = "heading 2")

  questions <- unique(df$Question)
  for (q in questions) {
    q_rows <- df %>% filter(Question == q) %>% select(-Question)

    # Truncate long question text for the heading (full text is in the table data)
    q_heading <- if (nchar(q) > 120) paste0(substr(q, 1, 117), "...") else q

    doc <- doc %>%
      body_add_par(q_heading, style = "heading 3") %>%
      body_add_flextable(make_table(q_rows)) %>%
      body_add_par("", style = "Normal")
  }
  doc
}

# ── Build document ────────────────────────────────────────────────────────────
doc <- read_docx()

# Title
doc <- doc %>%
  body_add_par("Cannabis Screening Survey: Frequency Report", style = "heading 1") %>%
  body_add_par(paste("Generated:", format(Sys.Date(), "%B %d, %Y")), style = "Normal") %>%
  body_add_par("", style = "Normal")

# Sample summary
n_elig   <- if (!is.null(eligible))   nrow(eligible[!duplicated(eligible$Question), ]) else 0
n_inelig <- if (!is.null(ineligible)) nrow(ineligible[!duplicated(ineligible$Question), ]) else 0

doc <- body_add_par(doc,
  sprintf("Eligible respondents (full survey): %s  |  Ineligible respondents (screener + demographics): %s",
          ifelse(!is.null(eligible),   max(eligible$`N (answered)`,   na.rm = TRUE), "—"),
          ifelse(!is.null(ineligible), max(ineligible$`N (answered)`, na.rm = TRUE), "—")),
  style = "Normal")
doc <- body_add_par(doc, "", style = "Normal")

# Eligible frequencies
doc <- add_freq_section(doc, eligible,   "Eligible Respondents — Full Survey")

# Ineligible frequencies
doc <- add_freq_section(doc, ineligible, "Ineligible Respondents — Screener & Demographics")

# County
if (!is.null(county_freq) && nrow(county_freq) > 0) {
  doc <- doc %>%
    body_add_par("County of Practice", style = "heading 2") %>%
    body_add_flextable(make_county_table(county_freq)) %>%
    body_add_par("", style = "Normal")
}

# ── Save ──────────────────────────────────────────────────────────────────────
print(doc, target = OUTPUT_FILE)
cat("Done. Report saved to:", basename(OUTPUT_FILE), "\n")
