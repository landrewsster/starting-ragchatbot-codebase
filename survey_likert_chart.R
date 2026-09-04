# survey_likert_chart.R
#
# Produces diverging bar charts for Likert attitude, confidence, and knowledge
# questions from the cannabis screening survey.
#
# Output PNGs (all saved next to FREQ_FILE):
#   _likert_chart1_collapsed/uncollapsed — screening attitude items (3)
#   _likert_chart2_collapsed/uncollapsed — safety/belief items (5)
#   _confidence_collapsed/uncollapsed   — confidence items (3)
#   _knowledge_collapsed/uncollapsed    — knowledge items (3)
#
# Don't know bars appear on a separate scale at the right, separated from the
# main scale by a dashed vertical line (secondary x-axis at top shows 0–100%).
# Knowledge questions have no Don't know option — no secondary scale shown.
#
# Install packages once:
#   install.packages(c("readxl", "dplyr", "ggplot2", "stringr", "forcats"))

library(readxl)
library(dplyr)
library(ggplot2)
library(stringr)
library(forcats)

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE      <- "Z:/Data/MCH survey/Data analysis/output"
FREQ_FILE <- file.path(BASE, "MCHHealthcareProvide-DataSetForLauraAndNo_DATA_LABELS_2026-09-03_1016_EDITED_frequencies.xlsx")

# 1. Define the graphics folder path and automatically create it if missing
GRAPHICS_DIR <- file.path(BASE, "graphics")
if (!dir.exists(GRAPHICS_DIR)) {
  dir.create(GRAPHICS_DIR)
}

# 2. Update this function to route the new .png files directly into that folder
make_out  <- function(tag, suffix = NULL) {
  full_tag <- if (!is.null(suffix)) paste0(tag, "_", suffix) else tag
  file.path(GRAPHICS_DIR, paste0(sub("\\.xlsx$", "", basename(FREQ_FILE)), "_", full_tag, ".png"))
}

# ── Response level sets (ordered inner → outer from center) ───────────────────
LIKERT_POS <- c("Agree", "Strongly agree")
LIKERT_NEG <- c("Disagree", "Strongly disagree")
CONF_POS   <- c("Somewhat confident", "Very confident")
CONF_NEG   <- c("Not very confident", "Not at all confident")
KNOW_POS      <- c("Somewhat knowledgeable", "Very Knowledgeable")
KNOW_NEG      <- c("Not Very Knowledgeable", "Not at all knowledgeable")
OPINION_POS     <- c("Somewhat supportive", "Very supportive")
OPINION_NEG     <- c("Somewhat opposed", "Very opposed")
OPINION_NEUTRAL <- c("Neither supportive nor opposed")
DK            <- "Don't know"

ALL_RESP_LEVELS <- unique(c(LIKERT_POS, LIKERT_NEG,
                             CONF_POS, CONF_NEG,
                             KNOW_POS, KNOW_NEG,
                             OPINION_POS, OPINION_NEG, OPINION_NEUTRAL,
                             DK))

# ── Question patterns ─────────────────────────────────────────────────────────
CHART1_PATTERNS <- c(
  "accurately report",
  "routine toxicology screening",
  "clinicians should screen"
)
CHART2_PATTERNS <- c(
  "there is no safe level",
  "potential risks.+outweigh",
  "therapeutic reasons",
  "contraindication to breastfeeding"
)
CONF_PATTERNS    <- c("talk about (cannabis|tobacco|alcohol)")
KNOW_PATTERNS    <- c("patients planning a pregnancy",
                      "patients who are pregnant",
                      "patients who are breastfeeding")
OPINION_PATTERNS <- c("feel about the legalization", "legalization of.*cannabis in minnesota")
ALL_PATTERNS  <- c(CHART1_PATTERNS, CHART2_PATTERNS, CONF_PATTERNS, KNOW_PATTERNS, OPINION_PATTERNS)

# ── Short labels ──────────────────────────────────────────────────────────────
# Listed most-specific first; first match wins.
SHORT_LABELS <- list(
  # Attitude — screening (chart 1)
  "accurately report"                  = "Patients accurately report\ncannabis use",
  "routine toxicology"                 = "Routine toxicology screening\nis appropriate",
  "clinicians should screen"           = "Clinicians should screen\nfor cannabis use",
  # Attitude — safety/belief (chart 2)
  "no safe level.+pregnancy"           = "No safe level of cannabis\nduring pregnancy",
  "no safe level.+breastfeed"          = "No safe level of cannabis\nduring breastfeeding",
  "no safe level"                      = "No safe level of cannabis",
  "risks.+fetus|risks.+pregnant"       = "Risks outweigh medical needs\n(fetus / pregnant person)",
  "risks.+newborn|risks.+breastfeed"   = "Risks outweigh medical needs\n(newborn / breastfeeding person)",
  "risks.+outweigh"                    = "Risks outweigh medical needs",
  "therapeutic reasons"                = "Patients use cannabis\nfor therapeutic reasons",
  "contraindication"                   = "Cannabis is contraindicated\nfor breastfeeding",
  # Confidence
  "talk about cannabis"                = "Confidence talking\nabout cannabis use",
  "talk about tobacco"                 = "Confidence talking\nabout tobacco use",
  "talk about alcohol"                 = "Confidence talking\nabout alcohol use",
  # Knowledge (specific sub-question first, generic fallback last)
  "planning a pregnancy"               = "Knowledge: patients\nplanning a pregnancy",
  "who are pregnant"                   = "Knowledge: patients\nwho are pregnant",
  "who are breastfeeding"              = "Knowledge: patients\nwho are breastfeeding",
  "level of knowledge.*pregnant"       = "Knowledge: patients\nwho are pregnant",
  "level of knowledge.*breastfeed"     = "Knowledge: patients\nwho are breastfeeding",
  "level of knowledge"                 = "Knowledge: patients",
  # Opinion — legalization
  "legalization of medical cannabis"           = "Medical cannabis\nlegalization (MN)",
  "non.medical.*recreational|recreational.*cannabis in minnesota" = "Recreational cannabis\nlegalization (MN)",
  "legalization.*cannabis in minnesota"        = "Cannabis legalization (MN)"  # fallback
)

# ── Colors ────────────────────────────────────────────────────────────────────
ALL_COLORS <- c(
  # Likert — uncollapsed
  "Strongly agree"           = "#1a5c32",
  "Agree"                    = "#74c476",
  "Disagree"                 = "#f4a261",
  "Strongly disagree"        = "#c0392b",
  # Confidence — uncollapsed
  "Very confident"           = "#4a1463",
  "Somewhat confident"       = "#b07cc6",
  "Not very confident"       = "#f4a261",
  "Not at all confident"     = "#c0392b",
  "Not confident"            = "#c0392b",   # kept for backward compat
  # Knowledge — uncollapsed
  "Very Knowledgeable"       = "#084594",
  "Somewhat knowledgeable"   = "#6baed6",
  "Not Very Knowledgeable"   = "#f4a261",
  "Not at all knowledgeable" = "#c0392b",
  # Collapsed labels (must match pos_label / neg_label passed to build_plot_data)
  "Knowledgeable"            = "#2166ac",
  "Not knowledgeable"        = "#c0392b",
  "Confident"                = "#7b2d8b",
  # "Agree" / "Disagree" / "Not confident" already covered above
  # Opinion — uncollapsed
  "Very supportive"                = "#1a5c32",
  "Somewhat supportive"            = "#74c476",
  "Neither supportive nor opposed" = "#cccccc",
  "Somewhat opposed"               = "#f4a261",
  "Very opposed"                   = "#c0392b",
  # Opinion — collapsed
  "Supportive"                     = "#2d7d46",
  "Opposed"                        = "#c0392b",
  "Neither"                        = "#cccccc",
  # DK (shared)
  "Don't know"               = "#aaaaaa"
)

# ── Layout constants ──────────────────────────────────────────────────────────
DK_START <- 110   # primary-axis x where DK bars begin
DK_SEP   <- DK_START - 4
MIN_W_N  <- 10    # min bar width (% units) to print % inside bar
N_X_DK   <- DK_START + 104   # n= label x when DK scale present
N_X_NODK <- 103              # n= label x when no DK scale

# ── Load and filter ───────────────────────────────────────────────────────────
eligible <- read_excel(FREQ_FILE, sheet = "eligible")

is_q_in <- function(q, patterns) {
  any(sapply(patterns, function(p) str_detect(str_to_lower(q), regex(p, ignore_case = TRUE))))
}

# Case-insensitive response normalization: maps lower-cased response strings to
# the canonical ALL_RESP_LEVELS spelling, so minor capitalisation differences in
# the Excel don't silently drop rows.
resp_norm_lookup <- setNames(ALL_RESP_LEVELS, str_to_lower(ALL_RESP_LEVELS))

raw <- eligible %>%
  filter(sapply(Question, is_q_in, patterns = ALL_PATTERNS)) %>%
  mutate(
    Response = {
      lwr <- str_to_lower(Response)
      dplyr::if_else(lwr %in% names(resp_norm_lookup),
                     resp_norm_lookup[lwr], Response)
    },
    n       = suppressWarnings(as.numeric(n)),
    N_denom = suppressWarnings(as.numeric(`N (denominator)`))
  ) %>%
  filter(Response %in% ALL_RESP_LEVELS)

if (nrow(raw) == 0) stop("No rows found — check FREQ_FILE path.")

# Assign short labels
get_short_label <- function(q) {
  for (pat in names(SHORT_LABELS))
    if (str_detect(str_to_lower(q), regex(pat, ignore_case = TRUE)))
      return(SHORT_LABELS[[pat]])
  str_wrap(q, width = 35)
}
q_label_map <- raw %>%
  distinct(Question) %>%
  mutate(q_short = sapply(Question, get_short_label))

dupes <- q_label_map$q_short[duplicated(q_label_map$q_short)]
if (length(dupes) > 0) {
  message("Duplicate short labels — falling back to wrapped full text: ",
          paste(dupes, collapse = "; "))
  q_label_map <- q_label_map %>%
    mutate(q_short = ifelse(q_short %in% dupes, str_wrap(Question, width = 45), q_short))
}
raw <- raw %>% left_join(q_label_map, by = "Question")

raw1     <- raw %>% filter(sapply(Question, is_q_in, patterns = CHART1_PATTERNS))
raw2     <- raw %>% filter(sapply(Question, is_q_in, patterns = CHART2_PATTERNS))
raw_conf    <- raw %>% filter(sapply(Question, is_q_in, patterns = CONF_PATTERNS))
raw_know    <- raw %>% filter(sapply(Question, is_q_in, patterns = KNOW_PATTERNS))
raw_opinion <- raw %>% filter(sapply(Question, is_q_in, patterns = OPINION_PATTERNS))

cat("Chart 1 (attitude-screening)  :", n_distinct(raw1$Question),       "questions\n")
cat("Chart 2 (attitude-safety)     :", n_distinct(raw2$Question),       "questions\n")
cat("Confidence                    :", n_distinct(raw_conf$Question),    "questions\n")
cat("Knowledge                     :", n_distinct(raw_know$Question),    "questions\n")
cat("Opinion (legalization)        :", n_distinct(raw_opinion$Question), "questions\n")

# ── Helper: build plot data ───────────────────────────────────────────────────
# pos_levels     : ordered inner→outer (e.g. c("Agree", "Strongly agree"))
# neg_levels     : ordered inner→outer (e.g. c("Disagree", "Strongly disagree"))
# dk_levels      : don't-know level(s); pass character(0) if none
# neutral_levels : midpoint level(s) centered at x=0; pass character(0) if none
# pos_label / neg_label / neutral_label : labels used when collapsed = TRUE
build_plot_data <- function(df_in, collapsed,
                             pos_levels, neg_levels, dk_levels = DK,
                             neutral_levels = character(0),
                             pos_label = "Positive", neg_label = "Negative",
                             neutral_label = "Neither") {

  if (nrow(df_in) == 0) stop("build_plot_data: df_in is empty — check question patterns match the Excel file")

  has_dk      <- length(dk_levels)      > 0 && any(df_in$Response %in% dk_levels,      na.rm = TRUE)
  has_neutral <- length(neutral_levels) > 0 && any(df_in$Response %in% neutral_levels, na.rm = TRUE)
  # "Neither" pulls out to its own zone (like DK). When both are present, Neither
  # starts one zone-width after DK; when Neither-only, it uses the same start position.
  neither_start <- if (has_dk) DK_START + 110L else DK_START
  all_resp    <- c(pos_levels, neg_levels, dk_levels, neutral_levels)

  # Per-question denominator from the Excel N (denominator) column
  n_denom_map <- df_in %>%
    group_by(Question) %>%
    summarise(n_all = first(N_denom[!is.na(N_denom)]), .groups = "drop")

  # ---- COLLAPSED ----
  if (collapsed) {
    df <- df_in %>%
      filter(Response %in% all_resp) %>%
      mutate(category = case_when(
        Response %in% pos_levels     ~ pos_label,
        Response %in% neg_levels     ~ neg_label,
        Response %in% dk_levels      ~ "Don't know",
        Response %in% neutral_levels ~ neutral_label,
        TRUE ~ NA_character_
      )) %>%
      filter(!is.na(category)) %>%
      group_by(Question, q_short, category) %>%
      summarise(n_count = sum(n, na.rm = TRUE), .groups = "drop") %>%
      left_join(n_denom_map, by = "Question") %>%
      group_by(Question) %>%
      mutate(n_all = first(n_all), pct = n_count / n_all * 100) %>%
      ungroup()

    anchors <- df %>%
      group_by(Question, q_short) %>%
      summarise(
        pos_pct     = sum(pct[category == pos_label],     na.rm = TRUE),
        neg_pct     = sum(pct[category == neg_label],     na.rm = TRUE),
        dk_pct      = sum(pct[category == "Don't know"],  na.rm = TRUE),
        neutral_pct = sum(pct[category == neutral_label], na.rm = TRUE),
        n_all       = first(n_all),
        .groups = "drop"
      ) %>%
      arrange(pos_pct) %>%
      mutate(y = row_number())

    q_order <- anchors$q_short
    n_df    <- df %>% select(Question, category, n_count)

    plot_df <- bind_rows(
      anchors %>% transmute(Question, q_short, y, category = pos_label,
                            xmin = 0,
                            xmax = pos_pct),
      anchors %>% transmute(Question, q_short, y, category = neg_label,
                            xmin = -neg_pct,
                            xmax = 0)
    )
    if (has_neutral)
      plot_df <- bind_rows(plot_df,
        anchors %>% transmute(Question, q_short, y, category = neutral_label,
                              xmin = neither_start, xmax = neither_start + neutral_pct))
    if (has_dk)
      plot_df <- bind_rows(plot_df,
        anchors %>% transmute(Question, q_short, y, category = "Don't know",
                              xmin = DK_START, xmax = DK_START + dk_pct))

    plot_df <- plot_df %>% left_join(n_df, by = c("Question", "category"))

    factor_lvls <- c(neg_label,
                     if (has_neutral) neutral_label,
                     pos_label,
                     if (has_dk) "Don't know")

  # ---- UNCOLLAPSED ----
  } else {
    df <- df_in %>%
      filter(Response %in% all_resp) %>%
      mutate(category = Response) %>%
      group_by(Question, q_short, category) %>%
      summarise(n_count = sum(n, na.rm = TRUE), .groups = "drop") %>%
      left_join(n_denom_map, by = "Question") %>%
      group_by(Question) %>%
      mutate(n_all = first(n_all), pct = n_count / n_all * 100) %>%
      ungroup()

    q_pos_order <- df %>%
      filter(category %in% pos_levels) %>%
      group_by(Question, q_short) %>%
      summarise(pos_total = sum(pct, na.rm = TRUE), n_all = first(n_all), .groups = "drop") %>%
      arrange(pos_total) %>%
      mutate(y = row_number())

    if (nrow(q_pos_order) == 0) {
      stop(paste0(
        "build_plot_data: no positive-level responses found in data.\n",
        "  pos_levels expected : ", paste(pos_levels, collapse = ", "), "\n",
        "  Response values seen: ", paste(sort(unique(df$category)), collapse = ", ")
      ))
    }

    q_order <- q_pos_order$q_short

    all_bars <- vector("list", nrow(q_pos_order))
    for (i in seq_len(nrow(q_pos_order))) {
      qrow  <- q_pos_order[i, ]
      q_df  <- df %>% filter(Question == qrow$Question)
      y_val <- qrow$y

      get_pct <- function(lv) {
        r <- q_df %>% filter(category == lv) %>% pull(pct)
        if (length(r) == 0) 0 else r[1]
      }
      get_n <- function(lv) {
        r <- q_df %>% filter(category == lv) %>% pull(n_count)
        if (length(r) == 0) NA_real_ else r[1]
      }

      # Neither is pulled out to its own zone; pos/neg stacks stay anchored at zero
      neutral_offset <- 0

      pos_pcts <- sapply(pos_levels, get_pct)
      pos_cum  <- cumsum(c(neutral_offset, pos_pcts))
      neg_pcts <- sapply(neg_levels, get_pct)
      neg_cum  <- cumsum(c(neutral_offset, neg_pcts))

      bars <- tibble(
        Question = qrow$Question,
        q_short  = qrow$q_short,
        y        = y_val,
        category = c(pos_levels, neg_levels),
        xmin     = c(pos_cum[seq_along(pos_levels)],
                     -neg_cum[seq_along(neg_levels) + 1]),
        xmax     = c(pos_cum[seq_along(pos_levels) + 1],
                     -neg_cum[seq_along(neg_levels)]),
        n_count  = c(sapply(pos_levels, get_n), sapply(neg_levels, get_n))
      )

      if (has_neutral) {
        neutral_pct_val <- sum(sapply(neutral_levels, get_pct))
        neutral_n_val   <- sum(sapply(neutral_levels, get_n), na.rm = TRUE)
        bars <- bind_rows(bars, tibble(
          Question = qrow$Question, q_short = qrow$q_short, y = y_val,
          category = neutral_label,
          xmin = neither_start, xmax = neither_start + neutral_pct_val,
          n_count = neutral_n_val
        ))
      }
      if (has_dk) {
        dk_pct_val <- sum(sapply(dk_levels, get_pct))
        dk_n_val   <- sum(sapply(dk_levels, get_n), na.rm = TRUE)
        bars <- bind_rows(bars, tibble(
          Question = qrow$Question, q_short = qrow$q_short, y = y_val,
          category = "Don't know",
          xmin = DK_START, xmax = DK_START + dk_pct_val,
          n_count = dk_n_val
        ))
      }
      all_bars[[i]] <- bars
    }

    plot_df <- bind_rows(all_bars)
    anchors <- q_pos_order %>% rename(pos_pct = pos_total)

    factor_lvls <- c(rev(neg_levels),
                     if (has_neutral) neutral_label,
                     pos_levels,
                     if (has_dk) "Don't know")
  }

  plot_df <- plot_df %>%
    mutate(
      category  = factor(category, levels = factor_lvls),
      xcenter   = (xmin + xmax) / 2,
      bar_width = abs(xmax - xmin)
    )

  list(plot_df = plot_df, anchors = anchors, q_order = q_order, has_dk = has_dk,
       has_neither = has_neutral, neither_start = neither_start)
}

# ── Chart builder ─────────────────────────────────────────────────────────────
make_chart <- function(df_in, collapsed,
                        pos_levels, neg_levels, dk_levels = DK,
                        neutral_levels = character(0),
                        pos_label = "Positive", neg_label = "Negative",
                        neutral_label = "Neither",
                        pos_arrow = NULL, neg_arrow = NULL,
                        title = "Cannabis Screening Survey",
                        missing_note = NULL) {

  d       <- build_plot_data(df_in, collapsed, pos_levels, neg_levels,
                              dk_levels, neutral_levels,
                              pos_label, neg_label, neutral_label)
  plot_df <- d$plot_df
  anchors <- d$anchors
  q_order <- d$q_order
  has_dk        <- d$has_dk
  has_neither   <- d$has_neither
  neither_start <- d$neither_start
  neither_sep   <- neither_start - 4
  n_q     <- length(q_order)

  n_x_pos     <- if (has_dk || has_neither) N_X_DK else N_X_NODK
  x_right_lim <- if (has_dk || has_neither) N_X_DK + 28 else N_X_NODK + 22

  pos_ann <- if (!is.null(pos_arrow)) pos_arrow else paste0(pos_label, " →")
  neg_ann <- if (!is.null(neg_arrow)) neg_arrow else paste0("← ", neg_label)

  caption_text <- glue_collapse(c(
    if (collapsed && length(pos_levels) > 1) paste0(pos_label, " = ", paste(rev(pos_levels), collapse = " + ")),
    if (collapsed && length(neg_levels) > 1) paste0(neg_label, " = ", paste(rev(neg_levels), collapse = " + ")),
    "Percentages of all respondents.",
    if (!is.null(missing_note)) missing_note
  ), sep = "  ")

  p <- ggplot(plot_df) +
    geom_rect(aes(
      xmin = xmin, xmax = xmax,
      ymin = y - 0.38, ymax = y + 0.38,
      fill = category
    ), color = "white", linewidth = 0.3) +
    geom_text(
      data = plot_df %>% filter(bar_width >= MIN_W_N),
      aes(x = xcenter, y = y, label = paste0(round(bar_width), "%")),
      color = "white", size = 4.2, fontface = "bold"
    ) +
    geom_vline(xintercept = 0, color = "black", linewidth = 0.5) +
    geom_text(
      data = anchors,
      aes(x = n_x_pos, y = y, label = paste0("n=", n_all)),
      hjust = 0, size = 4, color = "gray40", fontface = "bold"
    ) +
    annotate("text", x = -78, y = n_q + 0.75, label = neg_ann,
             hjust = 0, size = 4, color = "gray40", fontface = "bold.italic") +
    annotate("text", x = 78,  y = n_q + 0.75, label = pos_ann,
             hjust = 1, size = 4, color = "gray40", fontface = "bold.italic") +
    scale_y_continuous(
      breaks = seq_along(q_order), labels = q_order,
      expand = c(0.1, 0.1)
    ) +
    scale_fill_manual(values = ALL_COLORS) +
    labs(title = NULL, x = NULL, y = NULL, fill = NULL, caption = caption_text) +
    theme_minimal(base_size = 13) +
    theme(
      legend.position    = "bottom",
      legend.key.size    = unit(0.55, "cm"),
      legend.text        = element_text(size = 12, face = "bold"),
      panel.grid.major.y = element_blank(),
      panel.grid.minor   = element_blank(),
      panel.grid.major.x = element_line(color = "gray90", linewidth = 0.3),
      axis.text.y        = element_text(size = 12, lineheight = 1.2, face = "bold"),
      axis.text.x.bottom = element_text(size = 12, face = "bold"),
      axis.text.x.top    = element_text(size = 11, color = "gray50"),
      plot.title            = element_text(face = "bold", size = 14),
      plot.caption          = element_text(size = 9, color = "gray60", hjust = 0.5),
      plot.caption.position = "plot",
      plot.margin           = margin(10, 20, 10, 5)
    ) +
    guides(fill = guide_legend(nrow = 1)) +
    coord_cartesian(clip = "off")

  # DK secondary scale (only when Don't know responses are present)
  if (has_dk) {
    p <- p +
      geom_vline(xintercept = DK_SEP, color = "gray60",
                 linewidth = 0.4, linetype = "dashed") +
      geom_vline(xintercept = DK_START + c(25, 50, 75, 100),
                 color = "gray88", linewidth = 0.3) +
      annotate("text", x = DK_START + 1, y = n_q + 0.75,
               label = "Don't know →", hjust = 0, size = 4,
               color = "gray40", fontface = "bold.italic") +
      scale_x_continuous(
        limits = c(-100, x_right_lim),
        breaks = c(-100, -75, -50, -25, 0, 25, 50, 75, 100),
        labels = function(x) paste0(abs(x), "%"),
        expand = c(0, 0),
        sec.axis = sec_axis(
          transform = ~ . - DK_START,
          breaks    = seq(0, 100, 25),
          labels    = paste0(seq(0, 100, 25), "%"),
          name      = NULL
        )
      )
  } else if (has_neither) {
    # Neither pulled out to its own zone with its own 0-100% scale at the top
    p <- p +
      geom_vline(xintercept = neither_sep, color = "gray60",
                 linewidth = 0.4, linetype = "dashed") +
      geom_vline(xintercept = neither_start + c(25, 50, 75, 100),
                 color = "gray88", linewidth = 0.3) +
      annotate("text", x = neither_start + 1, y = n_q + 0.75,
               label = paste0(neutral_label, " →"), hjust = 0, size = 4,
               color = "gray40", fontface = "bold.italic") +
      scale_x_continuous(
        limits = c(-100, x_right_lim),
        breaks = c(-100, -75, -50, -25, 0, 25, 50, 75, 100),
        labels = function(x) paste0(abs(x), "%"),
        expand = c(0, 0),
        sec.axis = sec_axis(
          transform = ~ . - neither_start,
          breaks    = seq(0, 100, 25),
          labels    = paste0(seq(0, 100, 25), "%"),
          name      = NULL
        )
      )
  } else {
    p <- p +
      scale_x_continuous(
        limits = c(-100, x_right_lim),
        breaks = c(-100, -75, -50, -25, 0, 25, 50, 75, 100),
        labels = function(x) paste0(abs(x), "%"),
        expand = c(0, 0)
      )
  }

  p
}

# ── Caption helper (avoids importing glue) ────────────────────────────────────
glue_collapse <- function(x, sep = "") paste(x[nchar(x) > 0], collapse = sep)

# ── Missing-response footnote helper ─────────────────────────────────────────
# valid_questions: character vector of exact Question strings for this chart
#   (pass unique(raw_X$Question) so only the chart's own questions are checked).
# Python writes an explicit "Missing (skipped)" row for every question with
# n_missing > 0, so those rows are the only reliable source of truth.
get_missing_caption <- function(df_all, valid_questions) {
  miss_rows <- df_all %>%
    filter(Question %in% valid_questions,
           Response == "Missing (skipped)") %>%
    mutate(n_miss  = suppressWarnings(as.numeric(n)),
           N_denom = suppressWarnings(as.numeric(`N (denominator)`))) %>%
    filter(!is.na(n_miss), n_miss > 0, !is.na(N_denom))
  if (nrow(miss_rows) == 0) return(NULL)
  miss_by_q <- miss_rows %>%
    group_by(Question) %>%
    summarise(n_miss = first(n_miss), N_denom = first(N_denom), .groups = "drop") %>%
    mutate(pct_miss = n_miss / N_denom * 100)
  max_pct <- max(miss_by_q$pct_miss, na.rm = TRUE)
  n_qs    <- nrow(miss_by_q)
  if (n_qs == 1) {
    paste0(round(max_pct, 0), "% missing (skipped; n=", miss_by_q$n_miss[1], ")")
  } else {
    paste0("Up to ", round(max_pct, 0), "% missing (skipped) across ", n_qs, " questions")
  }
}


# ── Save charts ───────────────────────────────────────────────────────────────
LIKERT_TITLE  <- "Attitudes Toward Cannabis Use During Pregnancy and Breastfeeding"
CONF_TITLE    <- "Confidence in Discussing Substance Use with Patients"
KNOW_TITLE    <- "Knowledge of Cannabis Health Risks by Patient Population"
OPINION_TITLE <- "Attitudes Toward Cannabis Legalization in Minnesota"

# Chart 1 — attitude screening items (3 questions)
ggsave(make_out("likert_chart1", "uncollapsed"),
       make_chart(raw1, FALSE, LIKERT_POS, LIKERT_NEG, DK,
                  pos_label = "Agree", neg_label = "Disagree",
                  title = LIKERT_TITLE,
                  missing_note = get_missing_caption(eligible, unique(raw1$Question))),
       width = 11, height = 4.5, dpi = 300, bg = "white")
cat("Saved:", basename(make_out("likert_chart1", "uncollapsed")), "\n")

# Chart 2 — attitude safety/belief items (5 questions)
ggsave(make_out("likert_chart2", "uncollapsed"),
       make_chart(raw2, FALSE, LIKERT_POS, LIKERT_NEG, DK,
                  pos_label = "Agree", neg_label = "Disagree",
                  title = LIKERT_TITLE,
                  missing_note = get_missing_caption(eligible, unique(raw2$Question))),
       width = 13, height = 7, dpi = 300, bg = "white")
cat("Saved:", basename(make_out("likert_chart2", "uncollapsed")), "\n")

# Confidence chart (3 questions, has Don't know)
ggsave(make_out("confidence", "uncollapsed"),
       make_chart(raw_conf, FALSE, CONF_POS, CONF_NEG, DK,
                  pos_label = "Confident", neg_label = "Not confident",
                  pos_arrow = "Confident →", neg_arrow = "← Not confident",
                  title = CONF_TITLE,
                  missing_note = get_missing_caption(eligible, unique(raw_conf$Question))),
       width = 12, height = 4.5, dpi = 300, bg = "white")
cat("Saved:", basename(make_out("confidence", "uncollapsed")), "\n")

# Knowledge chart (3 questions, no Don't know)
ggsave(make_out("knowledge", "uncollapsed"),
       make_chart(raw_know, FALSE, KNOW_POS, KNOW_NEG, character(0),
                  pos_label = "Knowledgeable", neg_label = "Not knowledgeable",
                  pos_arrow = "Knowledgeable →", neg_arrow = "← Not knowledgeable",
                  title = KNOW_TITLE,
                  missing_note = get_missing_caption(eligible, unique(raw_know$Question))),
       width = 11, height = 4.5, dpi = 300, bg = "white")
cat("Saved:", basename(make_out("knowledge", "uncollapsed")), "\n")

# Opinion chart (2 questions, no DK, neutral midpoint centered at zero)
ggsave(make_out("opinion", "uncollapsed"),
       make_chart(raw_opinion, FALSE, OPINION_POS, OPINION_NEG, character(0),
                  neutral_levels = OPINION_NEUTRAL,
                  pos_label = "Supportive", neg_label = "Opposed",
                  neutral_label = "Neither",
                  pos_arrow = "Supportive →", neg_arrow = "← Opposed",
                  title = OPINION_TITLE,
                  missing_note = get_missing_caption(eligible, unique(raw_opinion$Question))),
       width = 11, height = 3.5, dpi = 300, bg = "white")
cat("Saved:", basename(make_out("opinion", "uncollapsed")), "\n")
