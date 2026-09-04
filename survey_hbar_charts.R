# survey_hbar_charts.R
#
# Horizontal bar charts for select-all-that-apply questions from the cannabis
# screening survey.  Bars are sorted by descending percentage; labels show
# both % and n.  Denominator = respondents eligible for each question
# (branching-aware, as computed by survey_frequencies.py).
#
# Output PNGs saved next to FREQ_FILE:
#   _sata_q45_inform_patients.png   — info needed to counsel patients
#   _sata_q46_screen.png            — info needed to improve screening
#   _sata_q47_inform_interventions.png — info needed for positive-screen response
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
make_out  <- function(tag) {
  file.path(GRAPHICS_DIR, paste0(sub("\\.xlsx$", "", basename(FREQ_FILE)), "_", tag, ".png"))
}

# ── Chart specs ───────────────────────────────────────────────────────────────
# Each list entry defines one chart.
# pattern : regex matched against the Question column (case-insensitive)
# stem    : output filename tag
# height  : saved height in inches (widen if responses wrap a lot)
CHARTS <- list(
  list(
    pattern = "better inform patients who are pregnant or breastfeeding",
    stem    = "sata_q45_inform_patients",
    height  = 6,
    title   = "Most providers lack guidance on cannabis risks during pregnancy and breastfeeding"
  ),
  list(
    pattern = "better screen patients who are pregnant or breastfeeding for cannabis",
    stem    = "sata_q46_screen",
    height  = 5.5,
    title   = "Standardized screening protocols are the top gap in cannabis screening practice"
  ),
  list(
    pattern = "better inform interventions",
    stem    = "sata_q47_inform_interventions",
    height  = 5.5,
    title   = "Providers report needing clear clinical pathways following a positive screen"
  ),

  # ── Additional SATA questions ──────────────────────────────────────────────
  list(
    pattern = "types of clinical encounters.+screen patients for cannabis",
    stem    = "sata_q6_encounter_types",
    height  = 5,
    title   = NULL   # fill in after reviewing data
  ),
  list(
    pattern = "positive screen for cannabis use during pregnancy",
    stem    = "sata_q10_actions_pregnancy",
    height  = 6,
    title   = NULL
  ),
  list(
    pattern = "positive screen for cannabis use during breastfeeding",
    stem    = "sata_q11_actions_breastfeeding",
    height  = 6,
    title   = NULL
  ),
  list(
    pattern = "how is this information recorded within the electronic health record",
    stem    = "sata_q15_ehr_recording",
    height  = 5.5,
    title   = NULL
  ),
  list(
    pattern = "factors may influence your confidence level discussing",
    stem    = "sata_q40_confidence_factors",
    height  = 6,
    title   = NULL
  )
)

# ── Load data ─────────────────────────────────────────────────────────────────
eligible <- read_excel(FREQ_FILE, sheet = "eligible") %>%
  mutate(
    pct = suppressWarnings(as.numeric(`%`)),
    n   = suppressWarnings(as.numeric(n)),
    N   = suppressWarnings(as.numeric(`N (denominator)`))
  )

cat(sprintf("Loaded eligible sheet: %d rows\n", nrow(eligible)))

# ── Helpers ───────────────────────────────────────────────────────────────────

get_sata_data <- function(df, pattern) {
  df %>%
    filter(
      Type == "Select all that apply",
      str_detect(Question, regex(pattern, ignore_case = TRUE)),
      Response != "Missing (skipped)"
    )
}

get_missing_note <- function(df, pattern) {
  # Branching questions: Python adds an explicit "Missing (skipped)" row — use it directly
  # so the "not asked" count (total_N - branching_N) is never mistaken for true skips.
  miss_row <- df %>%
    filter(
      str_detect(Question, regex(pattern, ignore_case = TRUE)),
      Response == "Missing (skipped)"
    )
  if (nrow(miss_row) > 0) {
    n_miss  <- miss_row$n[1]
    N_denom <- miss_row$N[1]
    if (is.na(n_miss) || n_miss == 0 || is.na(N_denom)) return(NULL)
    return(paste0(round(n_miss / N_denom * 100, 0), "% missing (skipped; n=", n_miss, ")"))
  }
  # Non-branching: N_denom = n_answered; missing = total_N - N_denom
  total_N <- max(df$N, na.rm = TRUE)
  if (!is.finite(total_N) || total_N == 0) return(NULL)
  qs <- df %>%
    filter(str_detect(Question, regex(pattern, ignore_case = TRUE)))
  if (nrow(qs) == 0) return(NULL)
  N_denom <- qs$N[1]
  if (is.na(N_denom)) return(NULL)
  n_miss <- max(total_N - N_denom, 0L)
  if (n_miss == 0) return(NULL)
  paste0(round(n_miss / total_N * 100, 0), "% missing (skipped; n=", n_miss, ")")
}

make_hbar <- function(df, title = NULL, missing_note = NULL) {
  N_denom <- df$N[1]

  plot_df <- df %>%
    filter(!is.na(pct)) %>%
    mutate(
      resp_wrapped = str_wrap(Response, width = 42),
      resp_wrapped = fct_reorder(resp_wrapped, pct)   # ascending = highest bar on top
    )

  x_max <- max(plot_df$pct, na.rm = TRUE)

  caption_base <- paste0(
    "n = ", N_denom,
    "  |  Each bar shows % of respondents who selected that option;",
    " totals may exceed 100%"
  )
  caption_text <- if (!is.null(missing_note)) {
    paste0(caption_base, ".  ", missing_note, ".")
  } else {
    paste0(caption_base, ".")
  }

  ggplot(plot_df, aes(x = pct, y = resp_wrapped)) +
    geom_col(fill = "#4472C4", width = 0.65) +
    geom_text(
      aes(label = paste0(round(pct, 0), "%  (n=", n, ")")),
      hjust    = -0.08,
      size     = 4.5,
      color    = "gray25",
      fontface = "bold"
    ) +
    scale_x_continuous(
      limits = c(0, x_max * 1.45),
      labels = function(x) paste0(x, "%"),
      expand = c(0, 0)
    ) +
    labs(
      title   = NULL,
      x       = NULL,
      y       = NULL,
      caption = caption_text
    ) +
    theme_minimal(base_size = 13) +
    theme(
      plot.title            = element_text(size = 13, face = "bold", hjust = 0,
                                           margin = margin(b = 8)),
      axis.text.y           = element_text(size = 12, hjust = 1, face = "bold"),
      axis.text.x           = element_text(size = 12, color = "gray50", face = "bold"),
      panel.grid.major.y    = element_blank(),
      panel.grid.minor      = element_blank(),
      panel.grid.major.x    = element_line(color = "gray90", linewidth = 0.4),
      plot.caption          = element_text(size = 9, color = "gray60", hjust = 0.5),
      plot.caption.position = "plot",
      plot.margin           = margin(10, 30, 10, 10)
    )
}

# ── Build and save ────────────────────────────────────────────────────────────
for (spec in CHARTS) {
  raw <- get_sata_data(eligible, spec$pattern)

  if (nrow(raw) == 0) {
    cat(sprintf("WARNING: no data matched pattern '%s'\n", spec$pattern))
    next
  }

  cat(sprintf("Building chart: %s  (%d response options)\n",
              spec$stem, nrow(raw)))

  miss_note  <- get_missing_note(eligible, spec$pattern)
  title_text <- if (!is.null(spec$title)) spec$title else unique(raw$Question)[1]
  p <- make_hbar(raw, title = title_text, missing_note = miss_note)
  ggsave(make_out(spec$stem), p,
         width = 11, height = spec$height, dpi = 300, bg = "white")
  cat(sprintf("  Saved: %s\n", basename(make_out(spec$stem))))
}

cat("\nDone.\n")
