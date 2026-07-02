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

# ── Question classification ───────────────────────────────────────────────────
DEMO_PATTERNS <- c(
  "what is your profession",
  "do you primarily see pregnant",
  "how long have you been practicing",
  "what is your primary specialty",
  "what is your secondary or sub.specialty",
  "which of the following best describes your primary practice setting",
  "how would you describe the insurance status",
  "what is the county of your practice",
  "what is your gender",
  "what is your age",
  "what is your race/ethnicity"
)

SCREENER_PATTERNS <- c(
  "provide prenatal.*delivery.*postpartum|prenatal.*delivery.*postpartum",
  "how many days.*see patients|average week.*how many"
)

SCREENER_LABELS <- c(
  "As part of your practice, do you provide prenatal, delivery or postpartum care?",
  "In an average week, how many days do you see patients?"
)

is_complete_q  <- function(q) str_detect(q, regex("Complete\\?", ignore_case = TRUE))
is_screener_q  <- function(q) sapply(q, function(x) any(str_detect(str_to_lower(x), SCREENER_PATTERNS)))
is_demo_q      <- function(q) sapply(q, function(x) any(str_detect(str_to_lower(x), DEMO_PATTERNS)))
is_main_q      <- function(q) !is_complete_q(q) & !is_screener_q(q) & !is_demo_q(q)

# Normalize question text: strip trailing pandas-dedup suffixes (.1, .2 …)
norm_q <- function(q) str_trim(str_remove(q, "\\s*\\.\\d+\\s*$"))

# ── Flextable helpers ─────────────────────────────────────────────────────────
# Single-group table: n and % columns, N as footer
make_table <- function(df) {
  n_col  <- intersect(c("N (answered)", "N (denominator)"), names(df))
  n_val  <- if (length(n_col) > 0) unique(df[[n_col[1]]])[1] else NA
  n_note <- if (!is.na(n_val)) paste0(n_col[1], " = ", n_val) else NULL

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

# Two-group side-by-side table with optional footer
make_wide_table <- function(df, footer = NULL) {
  num_cols <- names(df)[sapply(df, is.numeric)]
  ft <- flextable(df) %>%
    theme_booktabs() %>%
    bold(part = "header") %>%
    fontsize(size = 10, part = "all") %>%
    font(fontname = "Calibri", part = "all") %>%
    align(j = num_cols, align = "right", part = "all") %>%
    autofit()
  if (!is.null(footer)) {
    ft <- ft %>%
      add_footer_lines(footer) %>%
      fontsize(size = 9, part = "footer") %>%
      italic(part = "footer") %>%
      align(align = "left", part = "footer")
  }
  ft
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

# ── Build side-by-side table for one question ─────────────────────────────────
side_by_side <- function(elig_df, inelig_df, pattern = NULL, exact_q = NULL) {
  pull <- function(df, suffix) {
    if (is.null(df) || nrow(df) == 0) return(NULL)
    rows <- if (!is.null(exact_q)) {
      df %>% filter(norm_q(Question) == norm_q(exact_q))
    } else {
      df %>% filter(str_detect(str_to_lower(Question), regex(pattern, ignore_case = TRUE)))
    }
    if (nrow(rows) == 0) return(NULL)
    n_col <- intersect(c("N (answered)", "N (denominator)"), names(rows))
    n_val <- if (length(n_col) > 0) max(rows[[n_col[1]]], na.rm = TRUE) else NA
    # Aggregate duplicate responses caused by REDCap arm duplication
    out <- rows %>%
      group_by(Response) %>%
      summarise(n = sum(n, na.rm = TRUE), .groups = "drop") %>%
      mutate(`%` = round(n / n_val * 100, 1)) %>%
      rename(!!paste0(suffix, " n") := n, !!paste0(suffix, " %") := `%`)
    attr(out, "n_val") <- n_val
    out
  }

  eq <- pull(elig_df,   "Eligible")
  iq <- pull(inelig_df, "Ineligible")

  if (is.null(eq) && is.null(iq)) return(NULL)

  all_resp <- unique(c(eq$Response, iq$Response))
  tbl <- data.frame(Response = all_resp, stringsAsFactors = FALSE)
  if (!is.null(eq)) tbl <- left_join(tbl, eq, by = "Response")
  if (!is.null(iq)) tbl <- left_join(tbl, iq, by = "Response")
  tbl <- tbl %>% mutate(across(where(is.numeric), ~ replace_na(., 0)))

  en  <- if (!is.null(eq)) attr(eq, "n_val") else NA
  in_ <- if (!is.null(iq)) attr(iq, "n_val") else NA
  parts <- c(
    if (!is.na(en))  paste0("Eligible N = ", en),
    if (!is.na(in_)) paste0("Ineligible N = ", in_)
  )
  footer <- if (length(parts) > 0) paste(parts, collapse = "   |   ") else NULL

  list(tbl = tbl, footer = footer)
}

# ── Section builders ──────────────────────────────────────────────────────────

# Complete? summary at top (side by side)
add_complete_summary <- function(doc, elig_df, inelig_df) {
  rows <- bind_rows(
    elig_df   %>% filter(is_complete_q(Question)) %>% mutate(Group = "Eligible"),
    inelig_df %>% filter(is_complete_q(Question)) %>% mutate(Group = "Ineligible")
  )
  if (nrow(rows) == 0) return(doc)

  tbl <- rows %>%
    select(Group, Question, Response, n) %>%
    pivot_wider(names_from = Group, values_from = n, values_fill = 0) %>%
    rename(Form = Question)

  doc %>%
    body_add_par("Survey Completion", style = "heading 2") %>%
    body_add_flextable(make_wide_table(tbl)) %>%
    body_add_par("Note: N may vary across questions due to optional or skipped items.", style = "Normal") %>%
    body_add_par("", style = "Normal")
}

# Screener questions — combined across all respondents (no eligibility stratification)
add_screener_section <- function(doc, elig_df, inelig_df) {
  doc <- body_add_par(doc, "Screener Questions", style = "heading 2")

  combine_screener <- function(elig_df, inelig_df, pattern) {
    get_rows <- function(df) {
      if (is.null(df) || nrow(df) == 0) return(NULL)
      df %>% filter(str_detect(str_to_lower(Question), regex(pattern, ignore_case = TRUE)))
    }
    eq <- get_rows(elig_df)
    iq <- get_rows(inelig_df)
    if (is.null(eq) && is.null(iq)) return(NULL)

    combined <- bind_rows(eq, iq) %>%
      group_by(Response) %>%
      summarise(n = sum(n, na.rm = TRUE), .groups = "drop")
    total_n <- sum(combined$n)
    combined %>%
      mutate(`%` = round(n / total_n * 100, 1)) %>%
      arrange(desc(n))
  }

  for (i in seq_along(SCREENER_PATTERNS)) {
    tbl <- combine_screener(elig_df, inelig_df, SCREENER_PATTERNS[i])
    if (!is.null(tbl)) {
      total_n <- sum(tbl$n)
      ft <- flextable(tbl) %>%
        theme_booktabs() %>%
        bold(part = "header") %>%
        fontsize(size = 10, part = "all") %>%
        font(fontname = "Calibri", part = "all") %>%
        align(j = c("n", "%"), align = "right", part = "all") %>%
        width(j = "Response", width = 3.5) %>%
        width(j = c("n", "%"), width = 0.7) %>%
        add_footer_lines(paste0("N (answered) = ", total_n)) %>%
        fontsize(size = 9, part = "footer") %>%
        italic(part = "footer") %>%
        align(align = "left", part = "footer")

      doc <- doc %>%
        body_add_par(SCREENER_LABELS[i], style = "heading 3") %>%
        body_add_flextable(ft) %>%
        body_add_par("", style = "Normal")
    }
  }
  doc
}

# Main survey questions — eligible only, no truncation of question text
add_main_section <- function(doc, elig_df, heading) {
  df <- elig_df %>% filter(is_main_q(Question))
  if (nrow(df) == 0) return(doc)

  doc <- body_add_par(doc, heading, style = "heading 2")

  for (q in unique(df$Question)) {
    q_rows    <- df %>% filter(Question == q)
    q_type    <- if ("Type" %in% names(q_rows)) unique(q_rows$Type)[1] else ""
    type_note <- if (isTRUE(q_type == "Select all that apply")) " (select all that apply)" else ""
    q_rows    <- q_rows %>% select(-Question, -any_of("Type"))

    doc <- doc %>%
      body_add_par(paste0(q, type_note), style = "heading 3") %>%
      body_add_flextable(make_table(q_rows)) %>%
      body_add_par("", style = "Normal")
  }
  doc
}

# Demographic questions side by side
add_demo_section <- function(doc, elig_df, inelig_df, heading) {
  # Collect all unique normalized demographic questions from either group
  all_qs <- unique(c(
    unique(elig_df$Question)[is_demo_q(unique(elig_df$Question))],
    unique(inelig_df$Question)[is_demo_q(unique(inelig_df$Question))]
  ))
  norm_qs <- unique(norm_q(all_qs))
  if (length(norm_qs) == 0) return(doc)

  doc <- body_add_par(doc, heading, style = "heading 2")

  for (nq in norm_qs) {
    # Match using normalized question text
    eq <- elig_df   %>% filter(norm_q(Question) == nq)
    iq <- inelig_df %>% filter(norm_q(Question) == nq)
    if (nrow(eq) == 0 && nrow(iq) == 0) next

    q_type    <- if ("Type" %in% names(eq) && nrow(eq) > 0) unique(eq$Type)[1] else
                 if ("Type" %in% names(iq) && nrow(iq) > 0) unique(iq$Type)[1] else ""
    type_note <- if (isTRUE(q_type == "Select all that apply")) " (select all that apply)" else ""

    res <- side_by_side(elig_df, inelig_df, exact_q = nq)
    if (is.null(res)) next

    doc <- doc %>%
      body_add_par(paste0(nq, type_note), style = "heading 3") %>%
      body_add_flextable(make_wide_table(res$tbl, res$footer)) %>%
      body_add_par("", style = "Normal")
  }
  doc
}

# ── Build document ────────────────────────────────────────────────────────────
doc <- read_docx()

doc <- doc %>%
  body_add_par("Cannabis Screening Survey: Frequency Report", style = "heading 1") %>%
  body_add_par(paste("Generated:", format(Sys.Date(), "%B %d, %Y")), style = "Normal") %>%
  body_add_par("", style = "Normal")

# Complete? summary — commented out until REDCap form structure is confirmed
# add_complete_summary(doc, eligible, ineligible)

# Simple top summary instead
n_elig_screener <- if (!is.null(eligible) && nrow(eligible) > 0) {
  sc_rows <- eligible %>%
    filter(str_detect(str_to_lower(Question), SCREENER_PATTERNS[1])) %>%
    filter(str_detect(str_to_lower(Response), "^yes"))
  if (nrow(sc_rows) > 0) sc_rows$n[1] else max(eligible$`N (answered)`, na.rm = TRUE)
} else NA

doc <- doc %>%
  body_add_par("Summary", style = "heading 2") %>%
  body_add_par(
    sprintf("Eligible respondents (full survey): n = %s",
            ifelse(is.na(n_elig_screener), "—", n_elig_screener)),
    style = "Normal") %>%
  body_add_par(
    sprintf("Ineligible respondents (screener + demographics only): n = %s (note: some may not have answered all screener questions)",
            ifelse(is.null(ineligible), "—", max(ineligible$`N (answered)`, na.rm = TRUE))),
    style = "Normal") %>%
  body_add_par("Note: N varies by question due to optional or skipped items.", style = "Normal") %>%
  body_add_par("", style = "Normal")

# 2. Screener questions (eligible + ineligible side by side)
doc <- add_screener_section(doc, eligible, ineligible)

# 3. Full survey — eligible respondents only
doc <- add_main_section(doc, eligible, "Full Survey — Eligible Respondents")

# 4. Demographic questions — eligible + ineligible side by side
doc <- add_demo_section(doc, eligible, ineligible, "Demographic Questions — All Respondents")

# 5. County
if (!is.null(county_freq) && nrow(county_freq) > 0) {
  doc <- doc %>%
    body_add_par("County of Practice", style = "heading 2") %>%
    body_add_flextable(make_county_table(county_freq)) %>%
    body_add_par("", style = "Normal")
}

# ── Save ──────────────────────────────────────────────────────────────────────
print(doc, target = OUTPUT_FILE)
cat("Done. Report saved to:", basename(OUTPUT_FILE), "\n")
