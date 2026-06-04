# survey_likert_chart.R
#
# Diverging bar chart for Likert-scale attitude questions.
# Agree (SA+A) on right, Disagree (SD+D) on left, Don't know as separate bar.
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
AGREE_LEVELS    <- c("Strongly agree", "Agree")
DISAGREE_LEVELS <- c("Strongly disagree", "Disagree")
DK_LEVEL        <- "Don't know"

LIKERT_PATTERNS <- c(
  "there is no safe level",
  "potential risks.+outweigh",
  "therapeutic reasons",
  "contraindication to breastfeeding",
  "accurately report",
  "routine toxicology screening",
  "clinicians should screen"
)

# Short labels — update if full question text differs from these guesses
SHORT_LABELS <- list(
  "no safe level"            = "No safe level of cannabis\nduring pregnancy",
  "risks.+outweigh"          = "Risks outweigh benefits\nfor therapeutic use",
  "therapeutic reasons"      = "Patients use cannabis\nfor therapeutic reasons",
  "contraindication"         = "Cannabis is contraindicated\nfor breastfeeding",
  "accurately report"        = "Patients accurately report\ncannabis use",
  "routine toxicology"       = "Routine toxicology screening\nis appropriate",
  "clinicians should screen" = "Clinicians should screen\nfor cannabis use"
)

COLORS <- c(
  "Agree"      = "#2d7d46",
  "Disagree"   = "#c0392b",
  "Don't know" = "#aaaaaa"
)

DK_GAP <- 4   # visual gap (percentage points) between Agree and DK bars

# ── Load and filter ───────────────────────────────────────────────────────────
eligible <- read_excel(FREQ_FILE, sheet = "eligible")

is_likert_q <- function(q) {
  any(sapply(LIKERT_PATTERNS,
             function(p) str_detect(str_to_lower(q), regex(p, ignore_case = TRUE))))
}

raw <- eligible %>%
  filter(sapply(Question, is_likert_q)) %>%
  filter(Response %in% c(AGREE_LEVELS, DISAGREE_LEVELS, DK_LEVEL)) %>%
  mutate(n = as.numeric(n))

if (nrow(raw) == 0) stop("No Likert rows found — check FREQ_FILE path and sheet.")

# ── Short labels ──────────────────────────────────────────────────────────────
get_short_label <- function(q) {
  for (pat in names(SHORT_LABELS))
    if (str_detect(str_to_lower(q), regex(pat, ignore_case = TRUE)))
      return(SHORT_LABELS[[pat]])
  str_wrap(q, width = 45)
}

q_label_map <- raw %>%
  distinct(Question) %>%
  mutate(q_short = sapply(Question, get_short_label))

raw <- raw %>% left_join(q_label_map, by = "Question")

# ── Collapse to Agree / Disagree / Don't know ─────────────────────────────────
df <- raw %>%
  mutate(category = case_when(
    Response %in% AGREE_LEVELS    ~ "Agree",
    Response %in% DISAGREE_LEVELS ~ "Disagree",
    TRUE                          ~ "Don't know"
  )) %>%
  group_by(Question, q_short, category) %>%
  summarise(n = sum(n, na.rm = TRUE), .groups = "drop")

# Percentages: Agree/Disagree out of non-DK respondents; DK out of all
df <- df %>%
  group_by(Question) %>%
  mutate(
    n_no_dk = sum(n[category != "Don't know"]),
    n_all   = sum(n),
    pct     = ifelse(category == "Don't know",
                     n / n_all   * 100,
                     n / n_no_dk * 100)
  ) %>%
  ungroup()

# ── Bar x-positions ───────────────────────────────────────────────────────────
# Disagree: xmin = -disagree_pct, xmax = 0
# Agree:    xmin = 0,             xmax = agree_pct
# DK:       xmin = agree_pct + DK_GAP, xmax = agree_pct + DK_GAP + dk_pct

anchors <- df %>%
  group_by(Question, q_short) %>%
  summarise(
    agree_pct    = sum(pct[category == "Agree"],      na.rm = TRUE),
    disagree_pct = sum(pct[category == "Disagree"],   na.rm = TRUE),
    dk_pct       = sum(pct[category == "Don't know"], na.rm = TRUE),
    n_no_dk      = first(n_no_dk),
    .groups = "drop"
  )

# ── Question order (highest % agree at top) ───────────────────────────────────
q_order <- anchors %>%
  arrange(agree_pct) %>%
  pull(q_short)

anchors <- anchors %>%
  mutate(q_short = factor(q_short, levels = q_order),
         y = as.numeric(q_short))

# ── Long format for geom_rect ─────────────────────────────────────────────────
plot_df <- bind_rows(
  anchors %>% transmute(q_short, y, category = "Agree",
                        xmin = 0, xmax = agree_pct, pct = agree_pct),
  anchors %>% transmute(q_short, y, category = "Disagree",
                        xmin = -disagree_pct, xmax = 0, pct = disagree_pct),
  anchors %>% transmute(q_short, y, category = "Don't know",
                        xmin = agree_pct + DK_GAP,
                        xmax = agree_pct + DK_GAP + dk_pct,
                        pct = dk_pct)
) %>%
  mutate(
    category  = factor(category, levels = c("Disagree", "Agree", "Don't know")),
    xcenter   = (xmin + xmax) / 2,
    bar_width = xmax - xmin
  )

max_x <- max(plot_df$xmax, na.rm = TRUE)

# ── Plot ──────────────────────────────────────────────────────────────────────
p <- ggplot(plot_df) +
  geom_rect(aes(
    xmin = xmin, xmax = xmax,
    ymin = y - 0.38, ymax = y + 0.38,
    fill = category
  ), color = "white", linewidth = 0.3) +
  # Percentage labels (only on bars wide enough)
  geom_text(
    data = plot_df %>% filter(bar_width >= 8),
    aes(x = xcenter, y = y, label = paste0(round(pct), "%")),
    color = "white", size = 3, fontface = "bold"
  ) +
  # Center line
  geom_vline(xintercept = 0, color = "black", linewidth = 0.5) +
  # n= labels at far right
  geom_text(
    data = anchors,
    aes(x = max_x + 6, y = y, label = paste0("n=", n_no_dk)),
    hjust = 0, size = 2.8, color = "gray40"
  ) +
  # Disagree / Agree direction labels
  annotate("text", x = -78, y = length(q_order) + 0.75,
           label = "← Disagree", hjust = 0, size = 3, color = "gray40", fontface = "italic") +
  annotate("text", x = 78, y = length(q_order) + 0.75,
           label = "Agree →", hjust = 1, size = 3, color = "gray40", fontface = "italic") +
  # DK label above DK bars
  annotate("text",
           x = mean(c(anchors$agree_pct[1] + DK_GAP,
                       anchors$agree_pct[1] + DK_GAP + anchors$dk_pct[1])),
           y = length(q_order) + 0.75,
           label = "Don't\nknow", hjust = 0.5, size = 2.8, color = "gray50", fontface = "italic") +
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
  scale_fill_manual(values = COLORS) +
  labs(
    title    = "Attitudes Toward Cannabis Use During Pregnancy and Breastfeeding",
    subtitle = "Eligible respondents (full survey completed)",
    x = NULL, y = NULL, fill = NULL,
    caption  = paste0(
      "Agree = Strongly agree + Agree; Disagree = Strongly disagree + Disagree. ",
      "Agree/Disagree % exclude Don't know from denominator. n = respondents excluding Don't know."
    )
  ) +
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
  guides(fill = guide_legend(nrow = 1))

# ── Save ──────────────────────────────────────────────────────────────────────
ggsave(OUT_FILE, p, width = 11, height = 5.5, dpi = 300, bg = "white")
cat("Saved:", basename(OUT_FILE), "\n")
