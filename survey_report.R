# survey_report.R
#
# Reads survey_frequencies.xlsx and writes a formatted Word report.
#
# Install packages once (paste into RStudio console):
#   install.packages(c("readxl", "officer", "flextable", "dplyr", "stringr", "tidyr"))
#
# Open in RStudio and click Source (Cmd+Shift+S)

library(readxl)
library(officer)
library(flextable)
library(dplyr)
library(stringr)
library(tidyr)

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE <- file.path(path.expand("~"), "Downloads", "CRC MDH Project", "MDH analysis")

FREQ_FILE   <- file.path(BASE, "MCHHealthcareProvide-DataSetForLauraAndNo_DATA_LABELS_2026-06-01_1832_frequencies.xlsx")
OUTPUT_FILE <- sub("\\.xlsx$", "_report.docx", FREQ_FILE)

cat("Input :", basename(FREQ_FILE), "\n")
cat("Output:", basename(OUTPUT_FILE), "\n")

# ── Read sheets ───────────────────────────────────────────────────────────────
available  <- excel_sheets(FREQ_FILE)
read_if    <- function(name) if (name %in% available) read_excel(FREQ_FILE, sheet = name) else NULL

eligible    <- read_if("eligible")
ineligible  <- read_if("ineligible")
county_freq <- read_if("county_freq")

# ── Separate Complete? rows from main frequency data ─────────────────────────
separate_complete <- function(df) {
  if (is.null(df) || nrow(df) == 0) return(list(main = df, complete = NULL))
  is_complete <- str_detect(df$Question, regex("Complete\\?", ignore_case = TRUE))
  list(main = df[!is_complete, ], complete = df[is_complete, ])
}

elig_split   <- separate_complete(eligible)
inelig_split <- separate_complete(ineligible)

# ── Screener patterns (questions shown to ALL respondents) ────────────────────
SCREENER_PATTERNS <- c(
  "provide prenatal.*delivery.*postpartum|prenatal.*delivery.*postpartum",
  "how many days.*see patients|average week.*how many"
)

build_screener_table <- function(elig_df, inelig_df, pattern) {
  pull_q <- function(df, suffix) {
    if (is.null(df) || nrow(df) == 0) return(data.frame())
    df %>%
      filter(str_detect(Question, regex(pattern, ignore_case = TRUE))) %>%
      select(Response, n, `%`) %>%
      rename(!!paste0("Eligible n") := n, !!paste0("Eligible %") := `%`)
  }
  pull_q_i <- function(df) {
    if (is.null(df) || nrow(df) == 0) return(data.frame())
    df %>%
      filter(str_detect(Question, regex(pattern, ignore_case = TRUE))) %>%
      select(Response, n, `%`) %>%
      rename(`Ineligible n` = n, `Ineligible %` = `%`)
  }

  eq <- pull_q(elig_df,   "Eligible")
  iq <- pull_q_i(inelig_df)

  if (nrow(eq) == 0 && nrow(iq) == 0) return(NULL)

  all_resp <- unique(c(eq$Response, iq$Response))
  data.frame(Response = all_resp, stringsAsFactors = FALSE) %>%
    left_join(eq, by = "Response") %>%
    left_join(iq, by = "Response") %>%
    mutate(across(where(is.numeric), ~ replace_na(., 0)))
}

# ── Flextable helpers ─────────────────────────────────────────────────────────
make_table <- function(df) {
  # Pull N for footer before removing the column
  n_col  <- intersect(c("N (answered)", "N (denominator)"), names(df))
  n_val  <- if (length(n_col) > 0) unique(df[[n_col[1]]])[1] else NA
  n_note <- if (length(n_col) > 0 && n_col[1] == "N (denominator)")
              paste0("N (denominator) = ", n_val)
            else if (!is.na(n_val))
              paste0("N (answered) = ", n_val)
            else NULL

  df <- df %>% select(-any_of(c("N (answered)", "N (denominator)")))

  ft <- flextable(df) %>%
    theme_booktabs() %>%
    bold(part = "header") %>%
    fontsize(size = 10, part = "all") %>%
    font(fontname = "Calibri", part = "all") %>%
    align(j = c("n", "%"), align = "right", part = "all") %>%
    width(j = "Response", width = 3.5) %>%
    width(j = c("n", "%"), width = 0.7)

  if (!is.null(n_note)) {
    ft <- ft %>%
      add_footer_lines(n_note) %>%
      fontsize(size = 9, part = "footer") %>%
      italic(part = "footer") %>%
      align(align = "left", part = "footer")
  }
  ft
}

make_wide_table <- function(df) {
  # For screener and Complete? side-by-side tables
  num_cols <- names(df)[sapply(df, is.numeric)]
  flextable(df) %>%
    theme_booktabs() %>%
    bold(part = "header") %>%
    fontsize(size = 10, part = "all") %>%
    font(fontname = "Calibri", part = "all") %>%
    align(j = num_cols, align = "right", part = "all") %>%
    autofit()
}

make_county_table <- function(df) {
  flextable(df) %>%
    theme_booktabs() %>%
    bold(part = "header") %>%
    bold(i = nrow(df), part = "body") %>%
    fontsize(size = 10, part = "all") %>%
    font(fontname = "Calibri", part = "all") %>%
    align(j = -1, align = "right", part = "all") %>%
    autofit()
}

# ── Add main frequency section (excludes Complete? and screener questions) ────
add_freq_section <- function(doc, df, heading) {
  if (is.null(df) || nrow(df) == 0) return(doc)

  # Exclude Complete? and screener questions (handled separately at top)
  screener_regex <- paste(SCREENER_PATTERNS, collapse = "|")
  df <- df %>%
    filter(!str_detect(Question, regex("Complete\\?",    ignore_case = TRUE))) %>%
    filter(!str_detect(Question, regex(screener_regex,   ignore_case = TRUE)))

  if (nrow(df) == 0) return(doc)
  doc <- body_add_par(doc, heading, style = "heading 2")

  for (q in unique(df$Question)) {
    q_rows  <- df %>% filter(Question == q)
    q_type  <- if ("Type" %in% names(q_rows)) unique(q_rows$Type)[1] else ""
    q_rows  <- q_rows %>% select(-Question, -any_of("Type"))

    q_short   <- if (nchar(q) > 120) paste0(substr(q, 1, 117), "...") else q
    type_note <- if (isTRUE(q_type == "Select all that apply")) " (select all that apply)" else ""

    doc <- doc %>%
      body_add_par(paste0(q_short, type_note), style = "heading 3") %>%
      body_add_flextable(make_table(q_rows)) %>%
      body_add_par("", style = "Normal")
  }
  doc
}

# ── Build Complete? summary table (eligible + ineligible side by side) ────────
build_complete_summary <- function(elig_complete, inelig_complete) {
  bind_rows(
    if (!is.null(elig_complete) && nrow(elig_complete) > 0)
      elig_complete %>% mutate(Group = "Eligible"),
    if (!is.null(inelig_complete) && nrow(inelig_complete) > 0)
      inelig_complete %>% mutate(Group = "Ineligible")
  ) %>%
    select(Group, Question, Response, n) %>%
    pivot_wider(names_from = Group, values_from = n, values_fill = 0) %>%
    rename(Form = Question)
}

# ── Build document ────────────────────────────────────────────────────────────
doc <- read_docx()

# ── Title & top summary ───────────────────────────────────────────────────────
doc <- doc %>%
  body_add_par("Cannabis Screening Survey: Frequency Report", style = "heading 1") %>%
  body_add_par(paste("Generated:", format(Sys.Date(), "%B %d, %Y")), style = "Normal") %>%
  body_add_par("", style = "Normal") %>%
  body_add_par("Survey Completion", style = "heading 2")

# Complete? summary table
complete_tbl <- build_complete_summary(elig_split$complete, inelig_split$complete)
if (!is.null(complete_tbl) && nrow(complete_tbl) > 0) {
  doc <- doc %>%
    body_add_flextable(make_wide_table(complete_tbl)) %>%
    body_add_par("Note: N may vary by question due to skipped or optional items.", style = "Normal") %>%
    body_add_par("", style = "Normal")
}

# ── Screener questions (eligible + ineligible side by side) ───────────────────
screener_labels <- c(
  "Do you provide prenatal, delivery or postpartum care?",
  "In an average week, how many days do you see patients?"
)

doc <- body_add_par(doc, "Screener Questions", style = "heading 2")

for (i in seq_along(SCREENER_PATTERNS)) {
  tbl <- build_screener_table(elig_split$main, inelig_split$main, SCREENER_PATTERNS[i])
  if (!is.null(tbl) && nrow(tbl) > 0) {
    doc <- doc %>%
      body_add_par(screener_labels[i], style = "heading 3") %>%
      body_add_flextable(make_wide_table(tbl)) %>%
      body_add_par("", style = "Normal")
  }
}

# ── Eligible frequencies ──────────────────────────────────────────────────────
doc <- add_freq_section(doc, elig_split$main,   "Eligible Respondents — Full Survey")

# ── Ineligible frequencies ────────────────────────────────────────────────────
doc <- add_freq_section(doc, inelig_split$main, "Ineligible Respondents — Screener & Demographics")

# ── County ────────────────────────────────────────────────────────────────────
if (!is.null(county_freq) && nrow(county_freq) > 0) {
  doc <- doc %>%
    body_add_par("County of Practice", style = "heading 2") %>%
    body_add_flextable(make_county_table(county_freq)) %>%
    body_add_par("", style = "Normal")
}

# ── Save ──────────────────────────────────────────────────────────────────────
print(doc, target = OUTPUT_FILE)
cat("Done. Report saved to:", basename(OUTPUT_FILE), "\n")
