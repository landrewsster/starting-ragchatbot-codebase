# survey_likert_chart.R
#
# Produces two diverging bar charts for Likert-scale attitude questions:
#   1. _likert_collapsed.png   — Agree (SA+A) vs Disagree (SD+D)
#   2. _likert_uncollapsed.png — all four levels shown separately
#
# Bar widths are proportional to %; bar labels show n (counts).
# Don't know shown as a separate gray bar to the right.
#
# Install packages once:
#   install.packages(c("readxl", "dplyr", "ggplot2", "stringr", "forcats"))

library(readxl)
library(dplyr)
library(ggplot2)
library(stringr)
library(forcats)

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE      <- file.path(path.expand("~"), "Downloads", "CRC MDH Project", "MDH analysis")
FREQ_FILE <- file.path(BASE, "MCHHealthcareProvide-DataSetForLauraAndNo_DATA_LABELS_2026-06-01_1832_EDITED_frequencies.xlsx")
OUT_COLLAPSED   <- sub("\\.xlsx$", "_likert_collapsed.png",   FREQ_FILE)
OUT_UNCOLLAPSED <- sub("\\.xlsx$", "_likert_uncollapsed.png", FREQ_FILE)

# ── Settings ──────────────────────────────────────────────────────────────────
AGREE_LEVELS    <- c("Strongly agree", "Agree")
DISAGREE_LEVELS <- c("Strongly disagree", "Disagree")
DK_LEVEL        <- "Don't know"
ALL_LEVELS      <- c(DISAGREE_LEVELS, AGREE_LEVELS, DK_LEVEL)

LIKERT_PATTERNS <- c(
  "there is no safe level",
  "potential risks.+outweigh",
  "therapeutic reasons",
  "contraindication to breastfeeding",
  "accurately report",
  "routine toxicology screening",
  "clinicians should screen"
)

SHORT_LABELS <- list(
  "no safe level"            = "No safe level of cannabis\nduring pregnancy",
  "risks.+outweigh"          = "Risks outweigh benefits\nfor therapeutic use",
  "therapeutic reasons"      = "Patients use cannabis\nfor therapeutic reasons",
  "contraindication"         = "Cannabis is contraindicated\nfor breastfeeding",
  "accurately report"        = "Patients accurately report\ncannabis use",
  "routine toxicology"       = "Routine toxicology screening\nis appropriate",
  "clinicians should screen" = "Clinicians should screen\nfor cannabis use"
)

COLORS_COLLAPSED <- c(
  "Agree"      = "#2d7d46",
  "Disagree"   = "#c0392b",
  "Don't know" = "#aaaaaa"
)

COLORS_UNCOLLAPSED <- c(
  "Strongly agree"    = "#1a5c32",
  "Agree"             = "#74c476",
  "Disagree"          = "#f4a261",
  "Strongly disagree" = "#c0392b",
  "Don't know"        = "#aaaaaa"
)

DK_GAP   <- 4    # gap (% units) between Agree bar and DK bar
MIN_W_N  <- 5    # minimum bar width (% units) to show n label

# ── Load and filter ───────────────────────────────────────────────────────────
eligible <- read_excel(FREQ_FILE, sheet = "eligible")

is_likert_q <- function(q) {
  any(sapply(LIKERT_PATTERNS,
             function(p) str_detect(str_to_lower(q), regex(p, ignore_case = TRUE))))
}

raw <- eligible %>%
  filter(sapply(Question, is_likert_q)) %>%
  filter(Response %in% ALL_LEVELS) %>%
  mutate(n = as.numeric(n))

if (nrow(raw) == 0) stop("No Likert rows found — check FREQ_FILE path.")

# Short labels
get_short_label <- function(q) {
  for (pat in names(SHORT_LABELS))
    if (str_detect(str_to_lower(q), regex(pat, ignore_case = TRUE)))
      return(SHORT_LABELS[[pat]])
  str_wrap(q, width = 45)
}
q_label_map <- raw %>%
  distinct(Question) %>%
  mutate(q_short = sapply(Question, get_short_label))

# If two questions got the same short label, fall back to wrapped full text
dupes <- q_label_map$q_short[duplicated(q_label_map$q_short)]
if (length(dupes) > 0) {
  message("Duplicate short labels detected — using wrapped full text for affected questions: ",
          paste(dupes, collapse = "; "))
  q_label_map <- q_label_map %>%
    mutate(q_short = ifelse(q_short %in% dupes, str_wrap(Question, width = 45), q_short))
}

raw <- raw %>% left_join(q_label_map, by = "Question")

# ── Helper: build plot data ───────────────────────────────────────────────────
# Returns a data frame with xmin/xmax/xcenter/bar_width/n_count per bar segment,
# plus an anchors data frame for the n= annotations.

build_plot_data <- function(collapsed) {

  if (collapsed) {
    # Collapse SA+A → Agree, SD+D → Disagree
    df <- raw %>%
      mutate(category = case_when(
        Response %in% AGREE_LEVELS    ~ "Agree",
        Response %in% DISAGREE_LEVELS ~ "Disagree",
        TRUE                          ~ "Don't know"
      )) %>%
      group_by(Question, q_short, category) %>%
      summarise(n_count = sum(n, na.rm = TRUE), .groups = "drop")

    df <- df %>%
      group_by(Question) %>%
      mutate(
        n_no_dk = sum(n_count[category != "Don't know"]),
        n_all   = sum(n_count),
        pct     = n_count / n_all * 100
      ) %>%
      ungroup()

    anchors <- df %>%
      group_by(Question, q_short) %>%
      summarise(
        agree_pct    = sum(pct[category == "Agree"],      na.rm = TRUE),
        disagree_pct = sum(pct[category == "Disagree"],   na.rm = TRUE),
        dk_pct       = sum(pct[category == "Don't know"], na.rm = TRUE),
        n_no_dk      = first(n_no_dk),
        n_all        = first(n_all),
        .groups = "drop"
      ) %>%
      arrange(agree_pct) %>%
      mutate(y = row_number())

    q_order <- anchors$q_short

    n_df <- df %>% select(Question, category, n_count)

    plot_df <- bind_rows(
      anchors %>% transmute(Question, q_short, y, category = "Agree",
                            xmin = 0, xmax = agree_pct),
      anchors %>% transmute(Question, q_short, y, category = "Disagree",
                            xmin = -disagree_pct, xmax = 0),
      anchors %>% transmute(Question, q_short, y, category = "Don't know",
                            xmin = agree_pct + DK_GAP,
                            xmax = agree_pct + DK_GAP + dk_pct)
    ) %>%
      left_join(n_df, by = c("Question", "category")) %>%
      mutate(
        category  = factor(category, levels = c("Disagree", "Agree", "Don't know")),
        xcenter   = (xmin + xmax) / 2,
        bar_width = xmax - xmin
      )

  } else {
    # Uncollapsed: SA and A on right, D and SD on left
    df <- raw %>%
      mutate(category = Response) %>%
      group_by(Question, q_short, category) %>%
      summarise(n_count = sum(n, na.rm = TRUE), .groups = "drop")

    df <- df %>%
      group_by(Question) %>%
      mutate(
        n_no_dk = sum(n_count[category != "Don't know"]),
        n_all   = sum(n_count),
        pct     = n_count / n_all * 100
      ) %>%
      ungroup()

    anchors <- df %>%
      group_by(Question, q_short) %>%
      summarise(
        sa_pct = sum(pct[category == "Strongly agree"],    na.rm = TRUE),
        a_pct  = sum(pct[category == "Agree"],             na.rm = TRUE),
        d_pct  = sum(pct[category == "Disagree"],          na.rm = TRUE),
        sd_pct = sum(pct[category == "Strongly disagree"], na.rm = TRUE),
        dk_pct = sum(pct[category == "Don't know"],        na.rm = TRUE),
        n_no_dk = first(n_no_dk),
        n_all   = first(n_all),
        .groups = "drop"
      ) %>%
      mutate(agree_total = a_pct + sa_pct) %>%
      arrange(agree_total) %>%
      mutate(y = row_number())

    q_order <- anchors$q_short

    n_df <- df %>% select(Question, category, n_count)

    plot_df <- bind_rows(
      anchors %>% transmute(Question, q_short, y, category = "Strongly agree",
                            xmin = a_pct, xmax = a_pct + sa_pct),
      anchors %>% transmute(Question, q_short, y, category = "Agree",
                            xmin = 0, xmax = a_pct),
      anchors %>% transmute(Question, q_short, y, category = "Disagree",
                            xmin = -d_pct, xmax = 0),
      anchors %>% transmute(Question, q_short, y, category = "Strongly disagree",
                            xmin = -(d_pct + sd_pct), xmax = -d_pct),
      anchors %>% transmute(Question, q_short, y, category = "Don't know",
                            xmin = agree_total + DK_GAP,
                            xmax = agree_total + DK_GAP + dk_pct)
    ) %>%
      left_join(n_df, by = c("Question", "category")) %>%
      mutate(
        category  = factor(category, levels = c(
          "Strongly disagree", "Disagree", "Agree", "Strongly agree", "Don't know")),
        xcenter   = (xmin + xmax) / 2,
        bar_width = xmax - xmin
      )
  }

  list(plot_df = plot_df, anchors = anchors, q_order = q_order)
}

# ── Chart builder ─────────────────────────────────────────────────────────────
make_chart <- function(collapsed) {
  d      <- build_plot_data(collapsed)
  plot_df <- d$plot_df
  anchors <- d$anchors
  q_order <- d$q_order
  colors  <- if (collapsed) COLORS_COLLAPSED else COLORS_UNCOLLAPSED
  max_x   <- max(plot_df$xmax, na.rm = TRUE)

  caption_text <- if (collapsed) {
    "Agree = Strongly agree + Agree; Disagree = Strongly disagree + Disagree. Percentages of all respondents including Don't know."
  } else {
    "Percentages of all respondents including Don't know."
  }

  ggplot(plot_df) +
    geom_rect(aes(
      xmin = xmin, xmax = xmax,
      ymin = y - 0.38, ymax = y + 0.38,
      fill = category
    ), color = "white", linewidth = 0.3) +
    # n labels inside bars wide enough to fit
    geom_text(
      data = plot_df %>% filter(bar_width >= MIN_W_N, !is.na(n_count)),
      aes(x = xcenter, y = y, label = n_count),
      color = "white", size = 3, fontface = "bold"
    ) +
    geom_vline(xintercept = 0, color = "black", linewidth = 0.5) +
    # n= total at right margin
    geom_text(
      data = anchors,
      aes(x = max_x + 6, y = y, label = paste0("n=", n_all)),
      hjust = 0, size = 2.8, color = "gray40"
    ) +
    annotate("text", x = -78, y = length(q_order) + 0.75,
             label = "← Disagree", hjust = 0, size = 3,
             color = "gray40", fontface = "italic") +
    annotate("text", x = 78, y = length(q_order) + 0.75,
             label = "Agree →", hjust = 1, size = 3,
             color = "gray40", fontface = "italic") +
    scale_x_continuous(
      limits = c(-100, max_x + 16),
      breaks = c(-100, -75, -50, -25, 0, 25, 50, 75, 100),
      labels = function(x) paste0(abs(x), "%"),
      expand = c(0, 0)
    ) +
    scale_y_continuous(
      breaks = seq_along(q_order),
      labels = q_order,
      expand = c(0.1, 0.1)
    ) +
    scale_fill_manual(values = colors) +
    labs(
      title    = "Attitudes Toward Cannabis Use During Pregnancy and Breastfeeding",
      subtitle = "Eligible respondents (full survey completed)",
      x = NULL, y = NULL, fill = NULL,
      caption = caption_text
    ) +
    theme_minimal(base_size = 11) +
    theme(
      legend.position    = "bottom",
      legend.key.size    = unit(0.45, "cm"),
      legend.text        = element_text(size = 9),
      panel.grid.major.y = element_blank(),
      panel.grid.minor   = element_blank(),
      panel.grid.major.x = element_line(color = "gray90", linewidth = 0.3),
      axis.text.y        = element_text(size = 8, lineheight = 1.15),
      axis.text.x        = element_text(size = 9),
      plot.title         = element_text(face = "bold", size = 12),
      plot.subtitle      = element_text(size = 10, color = "gray50"),
      plot.caption       = element_text(size = 8, color = "gray60"),
      plot.margin        = margin(10, 10, 10, 40)
    ) +
    guides(fill = guide_legend(nrow = 1))
}

# ── Save both charts ──────────────────────────────────────────────────────────
ggsave(OUT_COLLAPSED,   make_chart(collapsed = TRUE),  width = 13, height = 5.5, dpi = 300, bg = "white")
cat("Saved:", basename(OUT_COLLAPSED), "\n")

ggsave(OUT_UNCOLLAPSED, make_chart(collapsed = FALSE), width = 13, height = 5.5, dpi = 300, bg = "white")
cat("Saved:", basename(OUT_UNCOLLAPSED), "\n")
