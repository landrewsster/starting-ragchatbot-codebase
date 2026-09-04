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
OUTPUT_DIR <- r"(Z:\Data\MCH survey\Data analysis\output)"

# Complete-case file (denominator varies by question) — exclude *allcases* files
freq_candidates <- list.files(OUTPUT_DIR,
  pattern = "frequencies.*\\.xlsx$",
  full.names = TRUE, ignore.case = TRUE)
freq_candidates <- freq_candidates[!grepl("allcases", freq_candidates, ignore.case = TRUE)]
if (length(freq_candidates) == 0)
  stop("No *frequencies*.xlsx file found in: ", OUTPUT_DIR)
FREQ_FILE <- freq_candidates[which.max(file.mtime(freq_candidates))]

# All-cases file (fixed N=122, explicit Missing rows for every question)
allcases_candidates <- list.files(OUTPUT_DIR,
  pattern = "frequencies_allcases.*\\.xlsx$",
  full.names = TRUE, ignore.case = TRUE)
ALLCASES_FILE <- if (length(allcases_candidates) > 0)
  allcases_candidates[which.max(file.mtime(allcases_candidates))] else NULL

# Set TRUE to read eligible tables from the all-cases file (recommended:
# shows Missing rows where n>0, footnotes where n=0)
USE_ALLCASES <- TRUE

TABLE_FILE  <- if (USE_ALLCASES && !is.null(ALLCASES_FILE)) ALLCASES_FILE else FREQ_FILE
OUTPUT_FILE <- sub("\\.xlsx$", "_report.docx", FREQ_FILE)

cat("Frequencies (complete-case):", basename(FREQ_FILE), "\n")
cat("Table source               :", basename(TABLE_FILE), "\n")
cat("Output                     :", basename(OUTPUT_FILE), "\n")

# ── Read sheets ───────────────────────────────────────────────────────────────
available  <- excel_sheets(FREQ_FILE)
read_if    <- function(name) if (name %in% available) read_excel(FREQ_FILE, sheet = name) else NULL

# eligible sheet comes from TABLE_FILE (all-cases when USE_ALLCASES=TRUE)
eligible <- {
  tbl_sheets <- excel_sheets(TABLE_FILE)
  if ("eligible" %in% tbl_sheets) read_excel(TABLE_FILE, sheet = "eligible") else NULL
}
ineligible       <- read_if("ineligible")
screener_data    <- read_if("screener")
crosstab_metro     <- read_if("crosstab_metro")
crosstab_tenure    <- read_if("crosstab_tenure")
crosstab_no_screen  <- read_if("crosstab_no_screen")
no_screen_freetext  <- read_if("no_screen_freetext")
crosstab_elig      <- read_if("crosstab_eligibility")
county_freq      <- read_if("county_freq")
county_data      <- read_if("county")
completion_time  <- read_if("completion_time")
completion_summ  <- read_if("completion_summary")

# Remove verbatim Open text rows — shown inline in Excel but excluded from the report
filter_open_text <- function(df) {
  if (is.null(df) || !("Type" %in% names(df))) return(df)
  df %>% filter(Type != "Open text")
}
eligible   <- filter_open_text(eligible)
ineligible <- filter_open_text(ineligible)

# ── Question classification ───────────────────────────────────────────────────
DEMO_PATTERNS <- c(
  "what is your profession",
  "do you primarily see pregnant",
  "how long have you been practicing",
  "what is your primary specialty",
  "what is your secondary or sub.specialty",
  "which of the following best describes your primary practice setting",
  "how would you describe the insurance status",
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

is_complete_q  <- function(q) str_detect(q, regex("Complete\\??\\s*$|_complete\\s*$|Complete\\?", ignore_case = TRUE))
is_screener_q  <- function(q) sapply(q, function(x) any(str_detect(str_to_lower(x), SCREENER_PATTERNS)))
is_demo_q      <- function(q) sapply(q, function(x) any(str_detect(str_to_lower(x), DEMO_PATTERNS)))
is_main_q      <- function(q) !is_complete_q(q) & !is_screener_q(q) & !is_demo_q(q)

# Normalize question text: strip trailing pandas-dedup suffixes (.1, .2 …)
norm_q <- function(q) str_trim(str_remove(q, "\\s*\\.\\d+\\s*$"))

# ── Ordered scale definitions (survey presentation order) ─────────────────────
SCALE_AGREE      <- c("Strongly agree", "Agree", "Disagree", "Strongly disagree", "Don't know")
SCALE_SUPPORT    <- c("Very supportive", "Somewhat supportive", "Neither supportive nor opposed",
                      "Somewhat opposed", "Very opposed")
SCALE_KNOWLEDGE  <- c("Very knowledgeable", "Somewhat knowledgeable",
                      "Not very knowledgeable", "Not at all knowledgeable")
SCALE_CONFIDENCE <- c("Very confident", "Somewhat confident",
                      "Not very confident", "Not at all confident")

ALL_ORDERED_SCALES <- list(SCALE_AGREE, SCALE_SUPPORT, SCALE_KNOWLEDGE, SCALE_CONFIDENCE)

# Return the matching ordered scale for df, or NULL if none applies.
detect_ordered_scale <- function(df) {
  if (!"Response" %in% names(df)) return(NULL)
  resp_low <- str_to_lower(str_squish(df$Response))
  for (scale in ALL_ORDERED_SCALES) {
    if (sum(resp_low %in% str_to_lower(scale)) >= 2) return(scale)
  }
  NULL
}

is_likert_df <- function(df) !is.null(detect_ordered_scale(df))

# Classify a column as "count" (n), "pct" (%), or "other"
.col_class <- function(col) {
  if (col == "n"  || grepl("^n \\(",  col)) return("count")
  if (col == "%"  || grepl("^% \\(",  col)) return("pct")
  "other"
}

# Reorder rows to match the detected ordered scale and zero-fill absent options.
# Total and Missing rows are preserved at the bottom in their original order.
enforce_likert_order <- function(df) {
  scale <- detect_ordered_scale(df)
  if (is.null(scale)) return(df)
  resp_lower  <- str_to_lower(str_squish(df$Response))
  scale_low   <- str_to_lower(scale)
  total_idx   <- which(resp_lower == "total")
  missing_idx <- which(str_detect(resp_lower, "^missing"))
  scale_idx   <- which(resp_lower %in% scale_low)
  other_idx   <- setdiff(seq_len(nrow(df)), c(scale_idx, total_idx, missing_idx))

  scale_rows <- lapply(scale, function(lv) {
    idx <- which(str_to_lower(str_squish(df$Response)) == str_to_lower(lv))
    if (length(idx) > 0) return(df[idx[1], , drop = FALSE])
    # Zero-filled row for this absent scale option
    new_row <- df[1, , drop = FALSE]
    new_row[["Response"]] <- lv
    for (col in setdiff(names(new_row), "Response")) {
      cls <- .col_class(col)
      if      (cls == "count") new_row[[col]] <- 0L
      else if (cls == "pct")   new_row[[col]] <- 0.0
      else if (is.numeric(new_row[[col]])) new_row[[col]] <- NA_real_
      else                                 new_row[[col]] <- NA_character_
    }
    new_row
  })

  bind_rows(c(
    scale_rows,
    if (length(other_idx)   > 0) list(df[other_idx,   ]) else list(),
    if (length(missing_idx) > 0) list(df[missing_idx, ]) else list(),
    if (length(total_idx)   > 0) list(df[total_idx,   ]) else list()
  ))
}

# ── Flextable helpers ─────────────────────────────────────────────────────────
# Single-group table: n and % columns, N as footer
make_table <- function(df) {
  # Strip "Not asked (branching rule)" rows before computing N or building table;
  # their count becomes a footer note so readers understand the branching.
  not_asked_mask <- str_detect(str_to_lower(str_squish(df$Response)), "^not asked")
  not_asked_rows <- df[not_asked_mask, ]
  df             <- df[!not_asked_mask, ]
  branching_note <- NULL
  if (nrow(not_asked_rows) > 0) {
    n_na <- suppressWarnings(as.numeric(not_asked_rows$n[1]))
    if (!is.na(n_na) && n_na > 0)
      branching_note <- paste0(
        n_na, " respondents were not routed to this question (branching rule); ",
        "percentages are based on those who were asked."
      )
  }

  # Find N: single-choice uses "N (answered)", checkbox uses "N (denominator)"
  n_col <- NA_character_
  n_val <- NA_real_
  for (.col in c("N (answered)", "N (denominator)")) {
    if (.col %in% names(df)) {
      .vals <- suppressWarnings(as.numeric(df[[.col]]))
      if (any(!is.na(.vals))) { n_col <- .col; n_val <- max(.vals, na.rm = TRUE); break }
    }
  }
  n_note <- if (!is.na(n_val)) paste0(n_col, " = ", as.integer(n_val)) else NULL

  df <- df %>% select(-any_of(c("N (answered)", "N (denominator)")))

  # Extract omnibus test info from first populated row, then drop stat columns
  stat_note <- NULL
  stat_cols_present <- intersect(c("p_value", "test", "stat_value", "stat_df"), names(df))
  if (length(stat_cols_present) > 0) {
    row1 <- df %>% filter(!is.na(p_value)) %>% slice(1)
    if (nrow(row1) > 0)
      stat_note <- fmt_stat_note(
        row1$p_value,
        if ("test"       %in% names(row1)) row1$test       else NA,
        if ("stat_value" %in% names(row1)) row1$stat_value else NA,
        if ("stat_df"    %in% names(row1)) row1$stat_df    else NA
      )
    df <- df %>% select(-any_of(stat_cols_present))
  }

  # Separate missing rows from substantive response rows
  missing_mask <- str_detect(str_to_lower(str_squish(df$Response)), "^missing")
  missing_rows <- df[missing_mask, ]
  df           <- df[!missing_mask, ]

  has_missing <- nrow(missing_rows) > 0 &&
    any(suppressWarnings(as.numeric(missing_rows$n)) > 0, na.rm = TRUE)

  no_missing_note <- if (!has_missing) "No missing responses for this question." else NULL

  # Sort non-Likert substantive rows by % descending; Likert order is handled below
  if (!is_likert_df(df) && "%" %in% names(df))
    df <- df %>% arrange(desc(suppressWarnings(as.numeric(.data[["%" ]]))))

  if (has_missing) df <- bind_rows(df, missing_rows)

  df <- enforce_likert_order(df)
  num_cols <- intersect(c("n", "%"), names(df))

  ft <- flextable(df) %>%
    theme_booktabs() %>%
    bold(part = "header") %>%
    fontsize(size = 10, part = "all") %>%
    font(fontname = "Calibri", part = "all") %>%
    colformat_num(na_str  = "") %>%
    colformat_char(na_str = "") %>%
    align(j = num_cols,              align = "right", part = "all") %>%
    align(j = setdiff(names(df), num_cols), align = "left", part = "all") %>%
    width(j = "Response", width = 3.5) %>%
    width(j = num_cols,   width = 0.7)

  if ("%" %in% names(df))
    ft <- colformat_double(ft, j = "%", digits = 1, na_str = "")
  if ("n" %in% names(df))
    ft <- colformat_double(ft, j = "n", digits = 0, na_str = "")

  footer_lines <- c(n_note, branching_note, stat_note, no_missing_note)
  footer_lines <- footer_lines[!sapply(footer_lines, is.null)]
  if (length(footer_lines) > 0) {
    for (fl in footer_lines)
      ft <- ft %>% add_footer_lines(fl)
    ft <- ft %>%
      fontsize(size = 9, part = "footer") %>%
      italic(part = "footer") %>%
      align(align = "left", part = "footer")
  }
  ft
}

# Two-group side-by-side table with optional footer
make_wide_table <- function(df, footer = NULL) {
  num_cols  <- names(df)[sapply(df, is.numeric)]
  resp_cols <- setdiff(names(df), num_cols)
  num_width  <- 0.65
  resp_width <- max(2.5, 6.5 - length(num_cols) * num_width)
  ft <- flextable(df) %>%
    theme_booktabs() %>%
    bold(part = "header") %>%
    fontsize(size = 10, part = "all") %>%
    font(fontname = "Calibri", part = "all") %>%
    align(j = num_cols,  align = "right", part = "all") %>%
    align(j = resp_cols, align = "left",  part = "all") %>%
    width(j = num_cols,  width = num_width) %>%
    width(j = resp_cols, width = resp_width)
  # footer may be a single string or a character vector; add each line separately
  footer_lines <- footer[!is.null(footer) & !is.na(footer) & nchar(footer) > 0]
  if (length(footer_lines) > 0) {
    for (fl in footer_lines)
      ft <- ft %>% add_footer_lines(fl)
    ft <- ft %>%
      fontsize(size = 9, part = "footer") %>%
      italic(part = "footer") %>%
      align(align = "left", part = "footer")
  }
  ft
}

fmt_p <- function(p) {
  ifelse(is.na(p), NA_character_,
         ifelse(as.numeric(p) < 0.001, "< 0.001", sprintf("%.3f", as.numeric(p))))
}

# Build a one-line omnibus test note for table footers.
# p_val, t_name, s_val, s_df may all be NA.
fmt_stat_note <- function(p_val, t_name, s_val = NA, s_df = NA) {
  p_val  <- suppressWarnings(as.numeric(p_val))
  s_val  <- suppressWarnings(as.numeric(s_val))
  s_df   <- suppressWarnings(as.integer(s_df))
  if (is.na(p_val) || is.na(t_name)) return(NULL)
  p_str <- if (p_val < 0.05) "p < .05" else "p > .05"
  if (isTRUE(str_detect(t_name, "χ²")) && !is.na(s_val) && !is.na(s_df))
    return(sprintf("Chi-square test statistic = %.2f (df = %d), %s", s_val, s_df, p_str))
  if (t_name == "Fisher")
    return(paste0("Fisher's exact test, ", p_str))
  if (isTRUE(str_detect(t_name, "Mann-Whitney")) && !is.na(s_val))
    return(sprintf("Mann-Whitney U = %.2f, %s", s_val, p_str))
  if (isTRUE(str_detect(t_name, "Wilcoxon")) && !is.na(s_val))
    return(sprintf("Wilcoxon signed rank test statistic = %.2f, %s", s_val, p_str))
  paste0(t_name, ": ", p_str)
}

# Cross-tab table for one question.
# Columns: Response | n/% per group | p (omnibus) | test | post_hoc_p (Holm-corrected)
# p_value:    omnibus chi-square or Fisher's exact (single-choice questions, first row only)
# post_hoc_p: Holm-corrected adjusted-residual p (single-choice, when omnibus p < .05)
#             OR Holm-corrected two-proportion Z-test p (select-all-that-apply, every row)
# Bottom row: "Total" with N1/N0 under n columns; footer: N = N1 + N0
make_crosstab_table <- function(df, label1, label0) {
  N1_col <- paste0("N (", label1, ")")
  N0_col <- paste0("N (", label0, ")")
  N1 <- if (N1_col %in% names(df)) as.integer(df[[N1_col]][1]) else NA_integer_
  N0 <- if (N0_col %in% names(df)) as.integer(df[[N0_col]][1]) else NA_integer_

  col_n1   <- paste0("n (", label1, ")")
  col_pct1 <- paste0("% (", label1, ")")
  col_n0   <- paste0("n (", label0, ")")
  col_pct0 <- paste0("% (", label0, ")")

  # Extract omnibus test note before narrowing columns
  stat_note <- NULL
  if ("p_value" %in% names(df)) {
    row1 <- df %>% filter(!is.na(p_value)) %>% slice(1)
    if (nrow(row1) > 0)
      stat_note <- fmt_stat_note(
        row1$p_value,
        if ("test"       %in% names(row1)) row1$test       else NA,
        if ("stat_value" %in% names(row1)) row1$stat_value else NA,
        if ("stat_df"    %in% names(row1)) row1$stat_df    else NA
      )
  }

  display <- df %>%
    select(any_of(c("Response", col_n1, col_pct1, col_n0, col_pct0, "post_hoc_p")))
  if ("post_hoc_p" %in% names(display))
    display <- display %>% mutate(post_hoc_p = fmt_p(post_hoc_p))

  display <- display %>% filter(!str_detect(str_to_lower(str_squish(Response)), "^missing"))

  # Sort non-Likert rows by first group's % descending; Likert order handled below
  if (!is_likert_df(display) && col_pct1 %in% names(display))
    display <- display %>% arrange(desc(suppressWarnings(as.numeric(.data[[col_pct1]]))))

  display <- enforce_likert_order(display)

  # Build Total row: N1/N0 under n columns, everything else blank
  total_row <- tibble(Response = "Total")
  for (cn in setdiff(names(display), "Response")) {
    if      (cn == col_n1)                        total_row[[cn]] <- as.numeric(N1)
    else if (cn == col_n0)                        total_row[[cn]] <- as.numeric(N0)
    else if (cn %in% c(col_pct1, col_pct0))       total_row[[cn]] <- NA_real_
    else                                          total_row[[cn]] <- NA_character_
  }
  display <- bind_rows(display, total_row)
  n_rows  <- nrow(display)

  num_cols  <- intersect(c(col_n1, col_pct1, col_n0, col_pct0), names(display))
  char_cols <- setdiff(names(display), num_cols)
  post_hoc_present <- intersect("post_hoc_p", names(display))

  n_footer <- paste0("N = ", N1 + N0)
  post_note <- if (length(post_hoc_present) > 0)
    "post_hoc_p: Holm-corrected per-option Z-test (SATA, always shown); adj. residual (single-choice, when omnibus p < .05). Likert/ordinal: Mann-Whitney U only — no post-hoc needed for a 2-group comparison." else NULL
  footer_lines <- c(n_footer, stat_note, post_note)
  footer_lines <- footer_lines[!sapply(footer_lines, is.null)]

  ft <- flextable(display) %>%
    theme_booktabs() %>%
    bold(part = "header") %>%
    bold(i = n_rows, part = "body") %>%
    fontsize(size = 10, part = "all") %>%
    font(fontname = "Calibri", part = "all") %>%
    colformat_num(na_str  = "") %>%
    colformat_char(na_str = "") %>%
    align(j = num_cols,  align = "right", part = "all") %>%
    align(j = char_cols, align = "left",  part = "all") %>%
    width(j = "Response",      width = 2.8) %>%
    width(j = num_cols,        width = 0.60) %>%
    width(j = post_hoc_present, width = 0.75)

  pct_display_cols <- intersect(c(col_pct1, col_pct0), names(display))
  n_display_cols   <- intersect(c(col_n1,   col_n0),   names(display))
  if (length(pct_display_cols) > 0)
    ft <- colformat_double(ft, j = pct_display_cols, digits = 1, na_str = "")
  if (length(n_display_cols) > 0)
    ft <- colformat_double(ft, j = n_display_cols, digits = 0, na_str = "")

  for (fl in footer_lines)
    ft <- ft %>% add_footer_lines(fl)
  ft %>%
    fontsize(size = 9, part = "footer") %>%
    italic(part = "footer") %>%
    align(align = "left", part = "footer")
}

# Loop through all questions in a cross-tab sheet and emit one table per question
add_crosstab_section <- function(doc, ct_df, label1, label0, heading) {
  if (is.null(ct_df) || nrow(ct_df) == 0) return(doc)
  ct_df <- ct_df %>% filter(is.na(Type) | Type != "Open text")
  if (nrow(ct_df) == 0) return(doc)

  doc <- body_add_par(doc, heading, style = "heading 2")

  for (q in unique(ct_df$Question)) {
    q_rows    <- ct_df %>% filter(Question == q)
    q_type    <- if ("Type" %in% names(q_rows)) unique(q_rows$Type)[1] else ""
    type_note <- if (isTRUE(q_type == "Select all that apply")) " (select all that apply)" else ""

    doc <- doc %>%
      body_add_par(paste0(q, type_note), style = "heading 3") %>%
      body_add_flextable(make_crosstab_table(q_rows, label1, label0)) %>%
      body_add_par("", style = "Normal")
  }
  doc
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
    # Find N: single-choice uses "N (answered)", checkbox uses "N (denominator)"
    n_val <- NA_real_
    for (.col in c("N (answered)", "N (denominator)")) {
      if (.col %in% names(rows)) {
        .vals <- suppressWarnings(as.numeric(rows[[.col]]))
        if (any(!is.na(.vals))) { n_val <- max(.vals, na.rm = TRUE); break }
      }
    }
    # Normalize whitespace in Response before aggregating to prevent spurious duplicates
    out <- rows %>%
      mutate(Response = str_squish(Response)) %>%
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
add_screener_section <- function(doc, screener_df) {
  if (is.null(screener_df) || nrow(screener_df) == 0) return(doc)
  doc <- body_add_par(doc, "Screener Questions", style = "heading 2")

  for (i in seq_along(SCREENER_PATTERNS)) {
    q_rows <- screener_df %>%
      filter(str_detect(str_to_lower(Question), regex(SCREENER_PATTERNS[i], ignore_case = TRUE))) %>%
      mutate(
        n = suppressWarnings(as.numeric(n)),
        `%` = suppressWarnings(as.numeric(`%`))
      )
    if (nrow(q_rows) == 0) next

    total_n <- suppressWarnings(as.numeric(
      q_rows[["N (denominator)"]][!is.na(q_rows[["N (denominator)"]])][1]
    ))
    if (is.na(total_n)) total_n <- sum(q_rows$n, na.rm = TRUE)

    tbl_data <- q_rows %>% select(Response, n, `%`)

    ft <- flextable(tbl_data) %>%
      theme_booktabs() %>%
      bold(part = "header") %>%
      fontsize(size = 10, part = "all") %>%
      font(fontname = "Calibri", part = "all") %>%
      align(j = c("n", "%"), align = "right", part = "all") %>%
      width(j = "Response", width = 3.5) %>%
      width(j = c("n", "%"), width = 0.7) %>%
      add_footer_lines(paste0("N (denominator) = ", total_n)) %>%
      fontsize(size = 9, part = "footer") %>%
      italic(part = "footer") %>%
      align(align = "left", part = "footer")

    doc <- doc %>%
      body_add_par(SCREENER_LABELS[i], style = "heading 3") %>%
      body_add_flextable(ft) %>%
      body_add_par("", style = "Normal")
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

add_elig_demo_section <- function(doc, elig_df, heading) {
  df <- elig_df %>% filter(is_demo_q(Question))
  if (nrow(df) == 0) return(doc)

  doc <- body_add_par(doc, heading, style = "heading 2")

  for (q in unique(norm_q(df$Question))) {
    q_rows    <- df %>% filter(norm_q(Question) == q)
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

# Single combined table for all demographic variables (eligible respondents).
# Bold shaded rows = question labels; response rows show n, %, and N (which varies by question).
make_elig_demo_combined_table <- function(elig_df, county_data = NULL) {
  df <- elig_df %>% filter(is_demo_q(Question))
  if (nrow(df) == 0) return(NULL)

  TENURE_PAT <- "how long have you been practicing"

  rows           <- list()
  header_indices <- integer(0)

  for (q in unique(norm_q(df$Question))) {
    # Raw tenure question is replaced by the derived binary rows appended below
    if (str_detect(str_to_lower(q), TENURE_PAT)) next

    q_rows <- df %>% filter(norm_q(Question) == q)

    n_val <- NA_real_
    for (.col in c("N (answered)", "N (denominator)")) {
      if (.col %in% names(q_rows)) {
        .vals <- suppressWarnings(as.numeric(q_rows[[.col]]))
        if (any(!is.na(.vals))) { n_val <- max(.vals, na.rm = TRUE); break }
      }
    }

    q_type    <- if ("Type" %in% names(q_rows)) unique(q_rows$Type)[1] else ""
    type_note <- if (isTRUE(q_type == "Select all that apply")) " (select all that apply)" else ""

    header_indices <- c(header_indices, length(rows) + 1L)
    rows[[length(rows) + 1L]] <- tibble(
      Response = paste0(q, type_note), n = NA_real_, `%` = NA_real_, N = NA_real_
    )

    for (i in seq_len(nrow(q_rows))) {
      rows[[length(rows) + 1L]] <- tibble(
        Response = q_rows$Response[i],
        n        = suppressWarnings(as.numeric(q_rows$n[i])),
        `%`      = suppressWarnings(as.numeric(q_rows$`%`[i])),
        N        = n_val
      )
    }
  }

  # Derived tenure binary (eligible only): replace raw question with ≥20 / <20 split
  tenure_rows <- df %>%
    filter(str_detect(str_to_lower(norm_q(Question)), TENURE_PAT),
           Response != "Missing (skipped)")
  if (nrow(tenure_rows) > 0) {
    n_high  <- sum(suppressWarnings(as.numeric(
                 tenure_rows$n[str_detect(str_to_lower(tenure_rows$Response), "20 or more")])),
                 na.rm = TRUE)
    n_low   <- sum(suppressWarnings(as.numeric(
                 tenure_rows$n[!str_detect(str_to_lower(tenure_rows$Response), "20 or more")])),
                 na.rm = TRUE)
    n_tot_t <- n_high + n_low
    if (n_tot_t > 0) {
      header_indices <- c(header_indices, length(rows) + 1L)
      rows[[length(rows) + 1L]] <- tibble(
        Response = "Tenure (years in practice)", n = NA_real_, `%` = NA_real_, N = NA_real_
      )
      rows[[length(rows) + 1L]] <- tibble(
        Response = "20 or more years (tenure = 1)",
        n = n_high, `%` = round(n_high / n_tot_t * 100, 1), N = as.numeric(n_tot_t)
      )
      rows[[length(rows) + 1L]] <- tibble(
        Response = "Less than 20 years (tenure = 0)",
        n = n_low, `%` = round(n_low / n_tot_t * 100, 1), N = as.numeric(n_tot_t)
      )
    }
  }

  # Derived metro binary (eligible only, requires county_data)
  if (!is.null(county_data) &&
      all(c("metro", "county_recoded", "group") %in% names(county_data))) {
    elig_metro <- county_data %>%
      filter(group == "eligible", county_recoded != "")
    n_metro   <- sum(as.integer(elig_metro$metro) == 1L)
    n_non     <- sum(as.integer(elig_metro$metro) == 0L)
    n_tot_m   <- n_metro + n_non
    if (n_tot_m > 0) {
      header_indices <- c(header_indices, length(rows) + 1L)
      rows[[length(rows) + 1L]] <- tibble(
        Response = "Metro (county of practice)", n = NA_real_, `%` = NA_real_, N = NA_real_
      )
      rows[[length(rows) + 1L]] <- tibble(
        Response = "Metro (7-county) (metro = 1)",
        n = as.numeric(n_metro), `%` = round(n_metro / n_tot_m * 100, 1), N = as.numeric(n_tot_m)
      )
      rows[[length(rows) + 1L]] <- tibble(
        Response = "Non-metro (metro = 0)",
        n = as.numeric(n_non), `%` = round(n_non / n_tot_m * 100, 1), N = as.numeric(n_tot_m)
      )
    }
  }

  combined <- bind_rows(rows)

  footer_lines <- c(
    "N = denominator for each question; may vary across questions due to branching or skipped items.",
    "Tenure = 1 if ≥20 years in practice; 0 otherwise.",
    "Metro = Hennepin, Washington, Carver, Scott, Ramsey, Anoka, or Dakota county."
  )

  flextable(combined) %>%
    theme_booktabs() %>%
    bold(part = "header") %>%
    bold(i = header_indices, part = "body") %>%
    bg(i = header_indices, bg = "#D9D9D9", part = "body") %>%
    fontsize(size = 10, part = "all") %>%
    font(fontname = "Calibri", part = "all") %>%
    colformat_num(na_str = "") %>%
    colformat_char(na_str = "") %>%
    align(j = c("n", "%", "N"), align = "right", part = "all") %>%
    align(j = "Response",       align = "left",  part = "all") %>%
    width(j = "Response",       width = 4.2) %>%
    width(j = c("n", "%", "N"), width = 0.65) %>%
    add_footer_lines(footer_lines) %>%
    fontsize(size = 9, part = "footer") %>%
    italic(part = "footer") %>%
    align(align = "left", part = "footer")
}

# Demographic questions side by side, with optional omnibus test results in footer.
# ct_elig: the crosstab_eligibility DataFrame (from Excel); NULL = no tests shown.
add_demo_section <- function(doc, elig_df, inelig_df, heading, ct_elig = NULL) {
  all_qs <- unique(c(
    unique(elig_df$Question)[is_demo_q(unique(elig_df$Question))],
    unique(inelig_df$Question)[is_demo_q(unique(inelig_df$Question))]
  ))
  norm_qs <- unique(norm_q(all_qs))
  if (length(norm_qs) == 0) return(doc)

  doc <- body_add_par(doc, heading, style = "heading 2")

  for (nq in norm_qs) {
    eq <- elig_df   %>% filter(norm_q(Question) == nq)
    iq <- inelig_df %>% filter(norm_q(Question) == nq)
    if (nrow(eq) == 0 && nrow(iq) == 0) next

    q_type    <- if ("Type" %in% names(eq) && nrow(eq) > 0) unique(eq$Type)[1] else
                 if ("Type" %in% names(iq) && nrow(iq) > 0) unique(iq$Type)[1] else ""
    type_note <- if (isTRUE(q_type == "Select all that apply")) " (select all that apply)" else ""

    res <- side_by_side(elig_df, inelig_df, exact_q = nq)
    if (is.null(res)) next

    # Look up omnibus test result for this question from the eligibility cross-tab
    stat_note <- NULL
    if (!is.null(ct_elig) && "Question" %in% names(ct_elig) && "p_value" %in% names(ct_elig)) {
      ct_q <- ct_elig %>% filter(norm_q(Question) == nq)
      row1 <- ct_q %>% filter(!is.na(p_value)) %>% slice(1)
      if (nrow(row1) > 0)
        stat_note <- fmt_stat_note(
          row1$p_value,
          if ("test"       %in% names(row1)) row1$test       else NA,
          if ("stat_value" %in% names(row1)) row1$stat_value else NA,
          if ("stat_df"    %in% names(row1)) row1$stat_df    else NA
        )
    }

    doc <- doc %>%
      body_add_par(paste0(nq, type_note), style = "heading 3") %>%
      body_add_flextable(make_wide_table(res$tbl, c(res$footer, stat_note))) %>%
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
  if (nrow(sc_rows) > 0) sc_rows$n[1] else max(eligible$`N (denominator)`, na.rm = TRUE)
} else NA

doc <- doc %>%
  body_add_par("Summary", style = "heading 2") %>%
  body_add_par(
    sprintf("Eligible respondents (full survey): n = %s",
            ifelse(is.na(n_elig_screener), "—", n_elig_screener)),
    style = "Normal") %>%
  body_add_par(
    sprintf("Ineligible respondents (screener + demographics only): n = %s (note: some may not have answered all screener questions)",
            ifelse(is.null(ineligible), "—", max(ineligible$`N (denominator)`, na.rm = TRUE))),
    style = "Normal") %>%
  body_add_par("Note: N varies by question due to optional or skipped items.", style = "Normal") %>%
  body_add_par("", style = "Normal")

# 2. Screener questions (eligible + ineligible combined)
doc <- add_screener_section(doc, screener_data)

# 3. Full survey — eligible respondents only
doc <- add_main_section(doc, eligible, "Full Survey — Eligible Respondents")

# 4. Demographic questions — eligible respondents only (single combined table)
ft_elig_demo <- make_elig_demo_combined_table(eligible, county_data = county_data)
if (!is.null(ft_elig_demo)) {
  doc <- doc %>%
    body_add_par("Demographic Questions — Eligible Respondents Only", style = "heading 2") %>%
    body_add_flextable(ft_elig_demo) %>%
    body_add_par("", style = "Normal")
}

# 4a. Demographic questions — eligible + ineligible side by side
doc <- add_demo_section(doc, eligible, ineligible, "Demographic Questions — All Respondents",
                        ct_elig = crosstab_elig)

# 4b. Tenure summary (derived binary from years of practice)
# Pull both eligible and ineligible counts from crosstab_eligibility, which
# already has the original tenure question answered by both groups.
if (!is.null(crosstab_elig) && nrow(crosstab_elig) > 0) {
  tenure_q <- crosstab_elig %>%
    filter(str_detect(str_to_lower(Question), "how long have you been practicing"))

  if (nrow(tenure_q) > 0 &&
      all(c("n (Eligible)", "n (Ineligible)") %in% names(tenure_q))) {

    high_rows <- tenure_q %>%
      filter(str_detect(str_to_lower(Response), "20 or more"),
             Response != "Missing (skipped)")
    low_rows  <- tenure_q %>%
      filter(!str_detect(str_to_lower(Response), "20 or more"),
             Response != "Missing (skipped)")

    e_high <- sum(suppressWarnings(as.integer(high_rows$`n (Eligible)`)),  na.rm = TRUE)
    e_low  <- sum(suppressWarnings(as.integer(low_rows$`n (Eligible)`)),   na.rm = TRUE)
    i_high <- sum(suppressWarnings(as.integer(high_rows$`n (Ineligible)`)), na.rm = TRUE)
    i_low  <- sum(suppressWarnings(as.integer(low_rows$`n (Ineligible)`)),  na.rm = TRUE)
    e_tot  <- e_high + e_low
    i_tot  <- i_high + i_low

    if (e_tot > 0 || i_tot > 0) {
      tenure_summary_tbl <- tibble(
        Response       = c("20 or more years (tenure = 1)", "Less than 20 years (tenure = 0)"),
        `Eligible n`   = as.integer(c(e_high, e_low)),
        `Eligible %`   = round(c(e_high, e_low) / max(e_tot, 1) * 100, 1),
        `Ineligible n` = as.integer(c(i_high, i_low)),
        `Ineligible %` = round(c(i_high, i_low) / max(i_tot, 1) * 100, 1),
      )
      footer_tenure <- paste0(
        if (e_tot > 0) paste0("Eligible N = ", e_tot) else NULL,
        if (e_tot > 0 && i_tot > 0) "   |   " else NULL,
        if (i_tot > 0) paste0("Ineligible N = ", i_tot) else NULL
      )
      doc <- doc %>%
        body_add_par("Tenure Group (Years in Practice)", style = "heading 2") %>%
        body_add_flextable(make_wide_table(tenure_summary_tbl,
          footer = c(footer_tenure,
                     "Tenure = 1 if 20 or more years in practice; tenure = 0 otherwise."))) %>%
        body_add_par("", style = "Normal")
    }
  }
}

# 4c. Metro group summary (derived from county of practice)
if (!is.null(county_data) &&
    all(c("metro", "county_recoded", "group") %in% names(county_data))) {

  valid_metro <- county_data %>% filter(county_recoded != "")
  e_metro <- sum(valid_metro$group == "eligible"   & as.integer(valid_metro$metro) == 1L)
  e_non   <- sum(valid_metro$group == "eligible"   & as.integer(valid_metro$metro) == 0L)
  i_metro <- sum(valid_metro$group == "ineligible" & as.integer(valid_metro$metro) == 1L)
  i_non   <- sum(valid_metro$group == "ineligible" & as.integer(valid_metro$metro) == 0L)
  e_tot_m <- e_metro + e_non
  i_tot_m <- i_metro + i_non

  if (e_tot_m > 0 || i_tot_m > 0) {
    metro_summary_tbl <- tibble(
      Response       = c("Metro (7-county) (metro = 1)", "Non-metro (metro = 0)"),
      `Eligible n`   = as.integer(c(e_metro, e_non)),
      `Eligible %`   = round(c(e_metro, e_non)   / max(e_tot_m, 1) * 100, 1),
      `Ineligible n` = as.integer(c(i_metro, i_non)),
      `Ineligible %` = round(c(i_metro, i_non) / max(i_tot_m, 1) * 100, 1),
    )
    footer_metro <- paste0(
      if (e_tot_m > 0) paste0("Eligible N = ", e_tot_m) else NULL,
      if (e_tot_m > 0 && i_tot_m > 0) "   |   " else NULL,
      if (i_tot_m > 0) paste0("Ineligible N = ", i_tot_m) else NULL
    )
    doc <- doc %>%
      body_add_par("Metro Group (County of Practice)", style = "heading 2") %>%
      body_add_flextable(make_wide_table(metro_summary_tbl,
        footer = c(footer_metro,
                   "Metro = Hennepin, Washington, Carver, Scott, Ramsey, Anoka, Dakota.",
                   "Only respondents who provided a valid county are included."))) %>%
      body_add_par("", style = "Normal")
  }
}

# 5. County
if (!is.null(county_freq) && nrow(county_freq) > 0) {
  doc <- doc %>%
    body_add_par("County of Practice", style = "heading 2") %>%
    body_add_flextable(make_county_table(county_freq)) %>%
    body_add_par("", style = "Normal")
}

# 5a. Metro vs Non-Metro
if (!is.null(county_data) &&
    all(c("metro", "county_recoded", "group") %in% names(county_data))) {

  valid    <- county_data %>% filter(county_recoded != "")
  elig_n   <- sum(valid$group == "eligible")
  inelig_n <- sum(valid$group == "ineligible")

  metro_tbl <- valid %>%
    mutate(Metro = if_else(as.integer(metro) == 1L,
                           "Metro (7-county)", "Non-metro")) %>%
    group_by(Metro) %>%
    summarise(
      `Eligible n`   = sum(group == "eligible"),
      `Ineligible n` = sum(group == "ineligible"),
      .groups = "drop"
    ) %>%
    mutate(
      `Eligible %`   = round(`Eligible n`   / elig_n   * 100, 1),
      `Ineligible %` = round(`Ineligible n` / inelig_n * 100, 1),
      `Total n`      = `Eligible n` + `Ineligible n`
    ) %>%
    select(Metro, `Eligible n`, `Eligible %`,
           `Ineligible n`, `Ineligible %`, `Total n`) %>%
    arrange(desc(`Total n`)) %>%
    bind_rows(tibble(
      Metro          = "TOTAL",
      `Eligible n`   = elig_n,
      `Eligible %`   = 100.0,
      `Ineligible n` = inelig_n,
      `Ineligible %` = 100.0,
      `Total n`      = elig_n + inelig_n
    ))

  doc <- doc %>%
    body_add_par("Metro vs Non-Metro", style = "heading 2") %>%
    body_add_flextable(
      make_wide_table(
        metro_tbl,
        footer = paste0(
          "Metro = Hennepin, Washington, Carver, Scott, Ramsey, Anoka, Dakota. ",
          "Respondents who provided a county only (N = ", elig_n + inelig_n, ")."
        )
      ) %>%
        bold(i = nrow(metro_tbl), part = "body")
    ) %>%
    body_add_par("", style = "Normal")
}

# 6. Completion time
if (!is.null(completion_time) && nrow(completion_time) > 0) {
  ct_stats <- completion_time %>%
    group_by(group) %>%
    summarise(
      n      = n(),
      Median = median(duration_minutes, na.rm = TRUE),
      Mean   = round(mean(duration_minutes, na.rm = TRUE), 1),
      Min    = min(duration_minutes, na.rm = TRUE),
      Max    = max(duration_minutes, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(group = str_to_title(group)) %>%
    rename(Group = group)

  ft_stats <- flextable(ct_stats) %>%
    theme_booktabs() %>%
    bold(part = "header") %>%
    fontsize(size = 10, part = "all") %>%
    font(fontname = "Calibri", part = "all") %>%
    align(j = -1, align = "right", part = "all") %>%
    add_footer_lines("All times in minutes. Median recommended over mean due to outliers (survey left open). Excludes respondents missing timestamps.") %>%
    fontsize(size = 9, part = "footer") %>%
    italic(part = "footer") %>%
    align(align = "left", part = "footer") %>%
    autofit()

  doc <- doc %>%
    body_add_par("Completion Time", style = "heading 2") %>%
    body_add_flextable(ft_stats) %>%
    body_add_par("", style = "Normal")

  if (!is.null(completion_summ) && nrow(completion_summ) > 0) {
    ft_buckets <- flextable(completion_summ) %>%
      theme_booktabs() %>%
      bold(part = "header") %>%
      fontsize(size = 10, part = "all") %>%
      font(fontname = "Calibri", part = "all") %>%
      align(j = -1, align = "right", part = "all") %>%
      autofit()

    doc <- doc %>%
      body_add_par("Distribution of Completion Times", style = "heading 3") %>%
      body_add_flextable(ft_buckets) %>%
      body_add_par("", style = "Normal")
  }
}

# 7. Cross-tabulations
doc <- add_crosstab_section(
  doc, crosstab_metro, "Metro", "Non-metro",
  "Cross-Tabulation by Metro / Non-Metro"
)
doc <- add_crosstab_section(
  doc, crosstab_tenure, "20+ years", "<20 years",
  "Cross-Tabulation by Tenure (Years in Practice)"
)
doc <- add_crosstab_section(
  doc, crosstab_no_screen,
  "No, don't screen any", "No, only some patients",
  "Cross-Tabulation: Q1.3.3 by Screening Group (Q1.3.1)"
)

# Q1.3.3 open-text responses by no-screen group
if (!is.null(no_screen_freetext) && nrow(no_screen_freetext) > 0 &&
    all(c("Group", "Response") %in% names(no_screen_freetext))) {
  ft_nsft <- flextable(no_screen_freetext %>% select(Group, Response)) %>%
    theme_booktabs() %>%
    bold(part = "header") %>%
    fontsize(size = 10, part = "all") %>%
    font(fontname = "Calibri", part = "all") %>%
    set_column_labels(Group = "Screening Group", Response = "Open-Text Response") %>%
    width(j = 1, width = 2) %>%
    width(j = 2, width = 4.5) %>%
    hrule(rule = "auto") %>%
    autofit()
  doc <- doc %>%
    body_add_par("Q1.3.3 Open-Text Responses by Screening Group", style = "heading 3") %>%
    body_add_flextable(ft_nsft) %>%
    body_add_par("", style = "Normal")
}

# ── Save ──────────────────────────────────────────────────────────────────────
print(doc, target = OUTPUT_FILE)
cat("Done. Report saved to:", basename(OUTPUT_FILE), "\n")
