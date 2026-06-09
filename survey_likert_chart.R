# survey_likert_chart.R
#
# Produces two pairs of diverging bar charts for Likert-scale attitude questions:
#   Chart 1 (3 statements): screening attitudes
#   Chart 2 (5 statements): safety / clinical belief items
#
# Each pair has a collapsed version (Agree vs Disagree) and an uncollapsed
# version (all four response levels shown separately).
#
# Don't know bars appear on a separate scale at the right, separated from the
# agree/disagree scale by a dashed vertical line. A secondary x-axis at the
# top shows 0–100% for the Don't know area.
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
make_out  <- function(tag, suffix) sub("\\.xlsx$", paste0("_", tag, "_", suffix, ".png"), FREQ_FILE)

# ── Settings ──────────────────────────────────────────────────────────────────
AGREE_LEVELS    <- c("Strongly agree", "Agree")
DISAGREE_LEVELS <- c("Strongly disagree", "Disagree")
DK_LEVEL        <- "Don't know"
ALL_LEVELS      <- c(DISAGREE_LEVELS, AGREE_LEVELS, DK_LEVEL)

# Questions for each chart
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
ALL_PATTERNS <- c(CHART1_PATTERNS, CHART2_PATTERNS)

SHORT_LABELS <- list(
  "no safe level.+pregnancy"           = "No safe level of cannabis\nduring pregnancy",
  "no safe level.+breastfeed"          = "No safe level of cannabis\nduring breastfeeding",
  "no safe level"                      = "No safe level of cannabis",
  "risks.+fetus|risks.+pregnant"       = "Risks outweigh medical needs\n(fetus / pregnant person)",
  "risks.+newborn|risks.+breastfeed"   = "Risks outweigh medical needs\n(newborn / breastfeeding person)",
  "risks.+outweigh"                    = "Risks outweigh medical needs",
  "therapeutic reasons"                = "Patients use cannabis\nfor therapeutic reasons",
  "contraindication"                   = "Cannabis is contraindicated\nfor breastfeeding",
  "accurately report"                  = "Patients accurately report\ncannabis use",
  "routine toxicology"                 = "Routine toxicology screening\nis appropriate",
  "clinicians should screen"           = "Clinicians should screen\nfor cannabis use"
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

# Don't know scale: fixed start position on primary x-axis
DK_START <- 110   # primary-axis x where DK bars begin (gap between 100 and 110)
DK_SEP   <- DK_START - 4   # dashed separator line x position
MIN_W_N  <- 5     # min bar width (% units) to print n inside bar
N_X      <- DK_START + 104 # fixed x for n= annotations (right of DK area)

# ── Load and filter ───────────────────────────────────────────────────────────
eligible <- read_excel(FREQ_FILE, sheet = "eligible")

is_q_in <- function(q, patterns) {
  any(sapply(patterns, function(p) str_detect(str_to_lower(q), regex(p, ignore_case = TRUE))))
}

raw <- eligible %>%
  filter(sapply(Question, is_q_in, patterns = ALL_PATTERNS)) %>%
  filter(Response %in% ALL_LEVELS) %>%
  mutate(n = as.numeric(n))

if (nrow(raw) == 0) stop("No Likert rows found — check FREQ_FILE path.")

# Short labels
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

raw1 <- raw %>% filter(sapply(Question, is_q_in, patterns = CHART1_PATTERNS))
raw2 <- raw %>% filter(sapply(Question, is_q_in, patterns = CHART2_PATTERNS))

cat("Chart 1 questions:", n_distinct(raw1$Question), "\n")
cat("Chart 2 questions:", n_distinct(raw2$Question), "\n")

# ── Helper: build plot data ───────────────────────────────────────────────────
build_plot_data <- function(df_in, collapsed) {

  if (collapsed) {
    df <- df_in %>%
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
        n_all = sum(n_count),
        pct   = n_count / n_all * 100
      ) %>%
      ungroup()

    anchors <- df %>%
      group_by(Question, q_short) %>%
      summarise(
        agree_pct    = sum(pct[category == "Agree"],      na.rm = TRUE),
        disagree_pct = sum(pct[category == "Disagree"],   na.rm = TRUE),
        dk_pct       = sum(pct[category == "Don't know"], na.rm = TRUE),
        n_all        = first(n_all),
        .groups = "drop"
      ) %>%
      arrange(agree_pct) %>%
      mutate(y = row_number())

    q_order <- anchors$q_short
    n_df    <- df %>% select(Question, category, n_count)

    plot_df <- bind_rows(
      anchors %>% transmute(Question, q_short, y, category = "Agree",
                            xmin = 0,            xmax = agree_pct),
      anchors %>% transmute(Question, q_short, y, category = "Disagree",
                            xmin = -disagree_pct, xmax = 0),
      anchors %>% transmute(Question, q_short, y, category = "Don't know",
                            xmin = DK_START,     xmax = DK_START + dk_pct)
    ) %>%
      left_join(n_df, by = c("Question", "category")) %>%
      mutate(
        category  = factor(category, levels = c("Disagree", "Agree", "Don't know")),
        xcenter   = (xmin + xmax) / 2,
        bar_width = xmax - xmin
      )

  } else {
    df <- df_in %>%
      mutate(category = Response) %>%
      group_by(Question, q_short, category) %>%
      summarise(n_count = sum(n, na.rm = TRUE), .groups = "drop")

    df <- df %>%
      group_by(Question) %>%
      mutate(
        n_all = sum(n_count),
        pct   = n_count / n_all * 100
      ) %>%
      ungroup()

    anchors <- df %>%
      group_by(Question, q_short) %>%
      summarise(
        sa_pct  = sum(pct[category == "Strongly agree"],    na.rm = TRUE),
        a_pct   = sum(pct[category == "Agree"],             na.rm = TRUE),
        d_pct   = sum(pct[category == "Disagree"],          na.rm = TRUE),
        sd_pct  = sum(pct[category == "Strongly disagree"], na.rm = TRUE),
        dk_pct  = sum(pct[category == "Don't know"],        na.rm = TRUE),
        n_all   = first(n_all),
        .groups = "drop"
      ) %>%
      mutate(agree_total = a_pct + sa_pct) %>%
      arrange(agree_total) %>%
      mutate(y = row_number())

    q_order <- anchors$q_short
    n_df    <- df %>% select(Question, category, n_count)

    plot_df <- bind_rows(
      anchors %>% transmute(Question, q_short, y, category = "Strongly agree",
                            xmin = a_pct,            xmax = a_pct + sa_pct),
      anchors %>% transmute(Question, q_short, y, category = "Agree",
                            xmin = 0,                xmax = a_pct),
      anchors %>% transmute(Question, q_short, y, category = "Disagree",
                            xmin = -d_pct,           xmax = 0),
      anchors %>% transmute(Question, q_short, y, category = "Strongly disagree",
                            xmin = -(d_pct + sd_pct), xmax = -d_pct),
      anchors %>% transmute(Question, q_short, y, category = "Don't know",
                            xmin = DK_START,          xmax = DK_START + dk_pct)
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
make_chart <- function(df_in, collapsed,
                       title = "Attitudes Toward Cannabis Use During Pregnancy and Breastfeeding") {
  d       <- build_plot_data(df_in, collapsed)
  plot_df <- d$plot_df
  anchors <- d$anchors
  q_order <- d$q_order
  colors  <- if (collapsed) COLORS_COLLAPSED else COLORS_UNCOLLAPSED
  n_q     <- length(q_order)

  caption_text <- if (collapsed) {
    "Agree = Strongly agree + Agree; Disagree = Strongly disagree + Disagree. Percentages of all respondents."
  } else {
    "Percentages of all respondents."
  }

  ggplot(plot_df) +
    # Agree / Disagree bars
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
    # Center (zero) line
    geom_vline(xintercept = 0, color = "black", linewidth = 0.5) +
    # Dashed separator between agree/disagree and Don't know scales
    geom_vline(xintercept = DK_SEP, color = "gray60",
               linewidth = 0.4, linetype = "dashed") +
    # Light grid lines in the Don't know area (25%, 50%, 75%, 100%)
    geom_vline(xintercept = DK_START + c(25, 50, 75, 100),
               color = "gray88", linewidth = 0.3) +
    # n= total labels at fixed right position
    geom_text(
      data = anchors,
      aes(x = N_X, y = y, label = paste0("n=", n_all)),
      hjust = 0, size = 2.8, color = "gray40"
    ) +
    # Direction annotations
    annotate("text", x = -78, y = n_q + 0.75,
             label = "← Disagree", hjust = 0, size = 3,
             color = "gray40", fontface = "italic") +
    annotate("text", x = 78,  y = n_q + 0.75,
             label = "Agree →", hjust = 1, size = 3,
             color = "gray40", fontface = "italic") +
    annotate("text", x = DK_START + 1, y = n_q + 0.75,
             label = "Don’t know →", hjust = 0, size = 3,
             color = "gray40", fontface = "italic") +
    # Primary x-axis (bottom): -100% to 100% for agree/disagree
    # Secondary x-axis (top):   0% to 100% for Don't know, via linear offset
    scale_x_continuous(
      limits = c(-100, N_X + 14),
      breaks = c(-100, -75, -50, -25, 0, 25, 50, 75, 100),
      labels = function(x) paste0(abs(x), "%"),
      expand = c(0, 0),
      sec.axis = sec_axis(
        transform = ~ . - DK_START,
        breaks    = seq(0, 100, 25),
        labels    = paste0(seq(0, 100, 25), "%"),
        name      = NULL
      )
    ) +
    scale_y_continuous(
      breaks = seq_along(q_order),
      labels = q_order,
      expand = c(0.1, 0.1)
    ) +
    scale_fill_manual(values = colors) +
    labs(
      title   = title,
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
      # Grid lines only in the primary agree/disagree area (suppressed in DK area
      # since we draw them manually above)
      panel.grid.major.x = element_line(color = "gray90", linewidth = 0.3),
      axis.text.y        = element_text(size = 7.5, lineheight = 1.2),
      axis.text.x.bottom = element_text(size = 9),
      axis.text.x.top    = element_text(size = 8, color = "gray50"),
      plot.title         = element_text(face = "bold", size = 12),
      plot.caption       = element_text(size = 8, color = "gray60"),
      plot.margin        = margin(10, 20, 10, 5)
    ) +
    guides(fill = guide_legend(nrow = 1))
}

# ── Save all four charts ──────────────────────────────────────────────────────
ggsave(make_out("likert_chart1", "collapsed"),
       make_chart(raw1, collapsed = TRUE),
       width = 11, height = 4.5, dpi = 300, bg = "white")
cat("Saved:", basename(make_out("likert_chart1", "collapsed")), "\n")

ggsave(make_out("likert_chart1", "uncollapsed"),
       make_chart(raw1, collapsed = FALSE),
       width = 11, height = 4.5, dpi = 300, bg = "white")
cat("Saved:", basename(make_out("likert_chart1", "uncollapsed")), "\n")

ggsave(make_out("likert_chart2", "collapsed"),
       make_chart(raw2, collapsed = TRUE),
       width = 13, height = 7, dpi = 300, bg = "white")
cat("Saved:", basename(make_out("likert_chart2", "collapsed")), "\n")

ggsave(make_out("likert_chart2", "uncollapsed"),
       make_chart(raw2, collapsed = FALSE),
       width = 13, height = 7, dpi = 300, bg = "white")
cat("Saved:", basename(make_out("likert_chart2", "uncollapsed")), "\n")
