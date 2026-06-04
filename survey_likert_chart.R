# survey_likert_chart.R
#
# Diverging stacked bar chart for Likert-scale attitude questions.
# Reads from the eligible sheet of the frequencies Excel file.
#
# Install packages once:
#   install.packages(c("readxl", "dplyr", "tidyr", "ggplot2", "stringr", "forcats"))

library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(stringr)
library(forcats)

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE      <- file.path(path.expand("~"), "Downloads", "CRC MDH Project", "MDH analysis")
FREQ_FILE <- file.path(BASE, "MCHHealthcareProvide-DataSetForLauraAndNo_DATA_LABELS_2026-06-01_1832_EDITED_frequencies.xlsx")
OUT_FILE  <- sub("\\.xlsx$", "_likert_chart.png", FREQ_FILE)

# ── Settings ──────────────────────────────────────────────────────────────────
LIKERT_LEVELS <- c(
  "Strongly agree", "Agree", "Neither agree nor disagree",
  "Disagree", "Strongly disagree", "Don't know"
)

# Patterns matching the Likert questions (mirrors survey_frequencies.py)
LIKERT_PATTERNS <- c(
  "there is no safe level",
  "potential risks.+outweigh",
  "therapeutic reasons",
  "contraindication to breastfeeding",
  "accurately report",
  "routine toxicology screening",
  "clinicians should screen"
)

# Optional: manually shorten question labels for the chart.
# Keys = substring of the full question (case-insensitive); values = short label.
# Leave this list empty to use auto-wrapped full question text.
SHORT_LABELS <- list(
  "no safe level"              = "No safe level of cannabis\nduring pregnancy",
  "risks.+outweigh"            = "Risks outweigh benefits\nfor therapeutic use",
  "therapeutic reasons"        = "Patients use cannabis\nfor therapeutic reasons",
  "contraindication"           = "Cannabis is contraindicated\nfor breastfeeding",
  "accurately report"          = "Patients accurately report\ncannabis use",
  "routine toxicology"         = "Routine toxicology screening\nis appropriate",
  "clinicians should screen"   = "Clinicians should screen\nfor cannabis use"
)

AGREE    <- c("Strongly agree", "Agree")
DISAGREE <- c("Strongly disagree", "Disagree")

COLORS <- c(
  "Strongly agree"             = "#1a6e3c",
  "Agree"                      = "#74c476",
  "Neither agree nor disagree" = "#d0d0d0",
  "Disagree"                   = "#f4a261",
  "Strongly disagree"          = "#c0392b",
  "Don't know"                 = "#aaaaaa"
)

# ── Load and filter ───────────────────────────────────────────────────────────
eligible <- read_excel(FREQ_FILE, sheet = "eligible")

is_likert_q <- function(q) {
  any(sapply(LIKERT_PATTERNS,
             function(p) str_detect(str_to_lower(q), regex(p, ignore_case = TRUE))))
}

df <- eligible %>%
  filter(sapply(Question, is_likert_q)) %>%
  filter(Response %in% LIKERT_LEVELS) %>%
  mutate(n = as.numeric(n))

if (nrow(df) == 0) stop("No Likert rows found — check that FREQ_FILE is up to date.")

# ── Short labels ──────────────────────────────────────────────────────────────
get_short_label <- function(q) {
  for (pattern in names(SHORT_LABELS)) {
    if (str_detect(str_to_lower(q), regex(pattern, ignore_case = TRUE)))
      return(SHORT_LABELS[[pattern]])
  }
  str_wrap(q, width = 45)  # fallback: auto-wrap full text
}

q_label_map <- df %>%
  distinct(Question) %>%
  mutate(q_short = sapply(Question, get_short_label))

df <- df %>% left_join(q_label_map, by = "Question")

# ── Percentages (excluding Don't know from denominator) ───────────────────────
df <- df %>%
  group_by(Question) %>%
  mutate(
    n_total    = sum(n[Response != "Don't know"], na.rm = TRUE),
    n_answered = sum(n, na.rm = TRUE),
    pct        = n / n_total * 100
  ) %>%
  ungroup()

# ── Diverging bar positions ───────────────────────────────────────────────────
# Neither agree nor disagree is centered at x = 0 (half on each side).
# Disagree categories extend left (negative), Agree categories extend right.
div_df <- df %>%
  filter(Response != "Don't know") %>%
  mutate(Response = factor(Response, levels = c(
    "Strongly disagree", "Disagree", "Neither agree nor disagree",
    "Agree", "Strongly agree"
  ))) %>%
  arrange(Question, Response) %>%
  group_by(Question) %>%
  mutate(
    neither_pct  = sum(pct[Response == "Neither agree nor disagree"], na.rm = TRUE),
    disagree_sum = sum(pct[Response %in% DISAGREE], na.rm = TRUE),
    offset       = -(disagree_sum + neither_pct / 2),
    cumulative   = cumsum(pct),
    xmax         = cumulative + offset,
    xmin         = lag(cumulative, default = 0) + offset,
    xcenter      = (xmin + xmax) / 2,
    bar_width    = xmax - xmin
  ) %>%
  ungroup()

# ── Question order (by % agree, ascending so highest is at top) ───────────────
q_order <- div_df %>%
  filter(Response %in% AGREE) %>%
  group_by(q_short) %>%
  summarise(pct_agree = sum(pct), .groups = "drop") %>%
  arrange(pct_agree) %>%
  pull(q_short)

div_df <- div_df %>%
  mutate(q_short = factor(q_short, levels = q_order))

# N labels shown at right margin
n_labels <- df %>%
  filter(Response == "Strongly agree") %>%   # one row per question
  distinct(q_short, n_total) %>%
  mutate(q_short = factor(q_short, levels = q_order))

# Don't know shown as separate annotation (% of all respondents)
dk_labels <- df %>%
  filter(Response == "Don't know") %>%
  mutate(pct_dk = n / n_answered * 100) %>%
  distinct(q_short, pct_dk) %>%
  mutate(
    q_short = factor(q_short, levels = q_order),
    label   = ifelse(pct_dk >= 1, paste0("DK: ", round(pct_dk), "%"), "")
  )

# ── Plot ──────────────────────────────────────────────────────────────────────
x_limit <- 105

p <- ggplot(div_df) +
  # Diverging bars
  geom_rect(aes(
    xmin = xmin, xmax = xmax,
    ymin = as.numeric(q_short) - 0.38,
    ymax = as.numeric(q_short) + 0.38,
    fill = Response
  ), color = "white", linewidth = 0.3) +
  # Percentage labels inside bars (only if bar is wide enough)
  geom_text(
    data = div_df %>% filter(bar_width >= 7),
    aes(x = xcenter, y = as.numeric(q_short),
        label = paste0(round(pct), "%")),
    size = 2.8, color = "white", fontface = "bold"
  ) +
  # Center line
  geom_vline(xintercept = 0, color = "black", linewidth = 0.4) +
  # N= labels at right
  geom_text(
    data = n_labels,
    aes(x = x_limit + 2, y = as.numeric(q_short),
        label = paste0("n=", n_total)),
    hjust = 0, size = 2.8, color = "gray40"
  ) +
  # Don't know labels
  geom_text(
    data = dk_labels %>% filter(label != ""),
    aes(x = x_limit + 2, y = as.numeric(q_short) - 0.55,
        label = label),
    hjust = 0, size = 2.4, color = "#888888", fontface = "italic"
  ) +
  # Scales
  scale_x_continuous(
    limits = c(-x_limit, x_limit + 14),
    breaks = seq(-100, 100, 25),
    labels = function(x) paste0(abs(x), "%"),
    expand = c(0, 0)
  ) +
  scale_y_continuous(
    breaks = seq_along(q_order),
    labels = q_order,
    expand = c(0.08, 0.08)
  ) +
  scale_fill_manual(
    values = COLORS,
    limits = c("Strongly disagree", "Disagree", "Neither agree nor disagree",
               "Agree", "Strongly agree")
  ) +
  # Labels
  labs(
    title    = "Attitudes Toward Cannabis Use During Pregnancy and Breastfeeding",
    subtitle = "Eligible respondents (full survey completed)",
    x        = NULL,
    y        = NULL,
    fill     = NULL,
    caption  = "Percentages exclude 'Don't know' responses from denominator. DK = Don't know."
  ) +
  # Annotations for sides
  annotate("text", x = -85, y = length(q_order) + 0.7,
           label = "← Disagree", hjust = 0, size = 3, color = "gray40", fontface = "italic") +
  annotate("text", x = 85,  y = length(q_order) + 0.7,
           label = "Agree →",   hjust = 1, size = 3, color = "gray40", fontface = "italic") +
  # Theme
  theme_minimal(base_size = 11) +
  theme(
    legend.position    = "bottom",
    legend.key.size    = unit(0.45, "cm"),
    legend.text        = element_text(size = 9),
    panel.grid.major.y = element_blank(),
    panel.grid.minor   = element_blank(),
    panel.grid.major.x = element_line(color = "gray90", linewidth = 0.3),
    axis.text.y        = element_text(size = 9, lineheight = 1.15),
    axis.text.x        = element_text(size = 9),
    plot.title         = element_text(face = "bold", size = 12),
    plot.subtitle      = element_text(size = 10, color = "gray50"),
    plot.caption       = element_text(size = 8, color = "gray60"),
    plot.margin        = margin(10, 10, 10, 10)
  ) +
  guides(fill = guide_legend(nrow = 1, reverse = TRUE))

# ── Save ──────────────────────────────────────────────────────────────────────
ggsave(OUT_FILE, p, width = 11, height = 5.5, dpi = 300, bg = "white")
cat("Saved:", basename(OUT_FILE), "\n")
