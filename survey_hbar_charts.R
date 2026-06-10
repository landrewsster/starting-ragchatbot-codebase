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
BASE      <- file.path(path.expand("~"), "Downloads", "CRC MDH Project", "MDH analysis")
FREQ_FILE <- file.path(BASE, "MCHHealthcareProvide-DataSetForLauraAndNo_DATA_LABELS_2026-06-01_1832_EDITED_frequencies.xlsx")
make_out  <- function(tag) sub("\\.xlsx$", paste0("_", tag, ".png"), FREQ_FILE)

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

make_hbar <- function(df, title = NULL) {
  N_denom <- df$N[1]

  plot_df <- df %>%
    filter(!is.na(pct)) %>%
    mutate(
      resp_wrapped = str_wrap(Response, width = 42),
      resp_wrapped = fct_reorder(resp_wrapped, pct)   # ascending = highest bar on top
    )

  x_max <- max(plot_df$pct, na.rm = TRUE)

  ggplot(plot_df, aes(x = pct, y = resp_wrapped)) +
    geom_col(fill = "#4472C4", width = 0.65) +
    geom_text(
      aes(label = paste0(round(pct, 0), "%  (n=", n, ")")),
      hjust  = -0.08,
      size   = 3.2,
      color  = "gray25"
    ) +
    scale_x_continuous(
      limits = c(0, x_max * 1.45),
      labels = function(x) paste0(x, "%"),
      expand = c(0, 0)
    ) +
    labs(
      title   = if (!is.null(title)) str_wrap(title, width = 90) else NULL,
      x       = NULL,
      y       = NULL,
      caption = paste0(
        "n = ", N_denom,
        "  |  Each bar shows % of respondents who selected that option;",
        " totals may exceed 100%"
      )
    ) +
    theme_minimal(base_size = 11) +
    theme(
      plot.title            = element_text(size = 11, face = "bold", hjust = 0,
                                           margin = margin(b = 8)),
      axis.text.y           = element_text(size = 9,  hjust = 1),
      axis.text.x           = element_text(size = 9,  color = "gray50"),
      panel.grid.major.y    = element_blank(),
      panel.grid.minor      = element_blank(),
      panel.grid.major.x    = element_line(color = "gray90", linewidth = 0.4),
      plot.caption          = element_text(size = 8, color = "gray60", hjust = 0.5),
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

  p <- make_hbar(raw, title = spec$title)
  ggsave(make_out(spec$stem), p,
         width = 11, height = spec$height, dpi = 300, bg = "white")
  cat(sprintf("  Saved: %s\n", basename(make_out(spec$stem))))
}

cat("\nDone.\n")
