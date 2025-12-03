# ============================================================
# Forest Plot for Exposure Windows (Corrected & Final)
# ============================================================
library(readxl)
library(forestplot)
library(dplyr)
library(stringr)
library(grid)

# ---------------- USER MUST EDIT ONLY THESE ----------------
file_path  <- "C:/Users/Bibin/Downloads/Forestplot_AQI_study_22_november.xlsx"
#sheet_name <- "complexXcontrols"
#sheet_name <- "all_cases"
#sheet_name <- "cyanotic"
sheet_name <- "acyanotic"
#sheet_name <- "simpleXcontrols"
output_SVG_file <- paste0("C:/Users/Bibin/Downloads/Mahima/Nov22_", sheet_name, ".svg")
# -----------------------------------------------------------

df <- read_excel(file_path, sheet = sheet_name)
colnames(df) <- trimws(colnames(df))

# Keep only rows with actual Adjusted OR (remove "_")
df <- df %>% filter(`ADJUSTED OR` != "_")

# --- Extract OR, lower CI, upper CI ---
df <- df %>%
  mutate(
    OR_val    = as.numeric(str_extract(`ADJUSTED OR`, "^[0-9.]+")),
    lower_CI  = as.numeric(str_extract(`ADJUSTED OR`, "(?<=\\().+?(?=-)")),
    upper_CI  = as.numeric(str_extract(`ADJUSTED OR`, "(?<=-).+?(?=\\))"))
  )

# --- Extract 4 values from case/control pairs ---
df <- df %>%
  mutate(
    bad_a  = as.numeric(sub("/.*", "", `CASE/CONTROLS(Bad outcome)`)),
    bad_b  = as.numeric(sub(".*/", "", `CASE/CONTROLS(Bad outcome)`)),
    good_c = as.numeric(sub("/.*", "", `CASE/CONTROLS(Good outcome)`)),
    good_d = as.numeric(sub(".*/", "", `CASE/CONTROLS(Good outcome)`)),
    total_n = bad_a + bad_b + good_c + good_d
  )

# --- Compute box size based on total sample size ---
max_boxsize <- 0.8
df <- df %>%
  mutate(
    boxsize_weight = total_n / max(total_n, na.rm = TRUE) * max_boxsize
  )

# --- Prepare labels ---
OR_CI <- sprintf("%.2f (%.2f–%.2f)", df$OR_val, df$lower_CI, df$upper_CI)

# Build the tabletext (first column header must match below checks)
tabletext <- cbind(
  c("Exposure Window", df$`EXPOSURE PERIOD`),
  c("Pollutant", df$POLLUTANTS),
  c("Exposure Level", df$`EXPOSURE LEVEL`),
  c("Adjusted OR (95% CI)", OR_CI)
)

# --- Group breaks (PRE–CONCEPTION–POST) ---
insert_after <- which(df$`EXPOSURE PERIOD` %in% c("PRECONCEPTION", "CONCEPTION"))

blank_row     <- c("", "", "", "")
tabletext_mod  <- tabletext[1, , drop = FALSE]
mean_mod       <- c(NA)
lower_mod      <- c(NA)
upper_mod      <- c(NA)
boxsize_mod    <- c(NA)

for (i in 2:nrow(tabletext)) {
  tabletext_mod  <- rbind(tabletext_mod, tabletext[i, , drop = FALSE])
  mean_mod       <- c(mean_mod, df$OR_val[i - 1])
  lower_mod      <- c(lower_mod, df$lower_CI[i - 1])
  upper_mod      <- c(upper_mod, df$upper_CI[i - 1])
  boxsize_mod    <- c(boxsize_mod, df$boxsize_weight[i - 1])
  
  if ((i - 1) %in% insert_after) {
    tabletext_mod  <- rbind(tabletext_mod, blank_row)
    mean_mod       <- c(mean_mod, NA)
    lower_mod      <- c(lower_mod, NA)
    upper_mod      <- c(upper_mod, NA)
    boxsize_mod    <- c(boxsize_mod, NA)
  }
}

# Remove fully empty accidental rows
keep_rows <- !(
  is.na(mean_mod) &
    is.na(lower_mod) &
    is.na(upper_mod) &
    is.na(boxsize_mod) &
    apply(tabletext_mod, 1, function(x) all(x == ""))
)

tabletext_mod <- tabletext_mod[keep_rows, , drop = FALSE]
mean_mod      <- mean_mod[keep_rows]
lower_mod     <- lower_mod[keep_rows]
upper_mod     <- upper_mod[keep_rows]
boxsize_mod   <- boxsize_mod[keep_rows]

# Add intentional trailing blank row (if you want)
tabletext_mod <- rbind(tabletext_mod, c("", "", "", ""))
mean_mod      <- c(mean_mod, NA)
lower_mod     <- c(lower_mod, NA)
upper_mod     <- c(upper_mod, NA)
boxsize_mod   <- c(boxsize_mod, NA)

# line heights / bottom padding
bottom_padding <- 0.8
lineheight_vec <- rep(unit(1, "cm"), length(mean_mod))
lineheight_vec[length(lineheight_vec)] <- unit(bottom_padding, "cm")


# ============================================================
# FINAL ROW-NUMBER BASED COLORING (USER CONTROLS EVERYTHING)
# ============================================================

n_rows <- nrow(tabletext_mod)

# default: all rows white
row_fill <- rep(list(gpar(fill = "white")), n_rows)

# ---------------- USER EDIT THESE ONLY ----------------

# Example:
blue_start   <- 2
blue_end     <- 5

yellow_start <- 6
yellow_end   <- 8

pink_start   <- 9
pink_end     <- 12

# -------------------------------------------------------

# Apply colors
row_fill[(blue_start-1):(blue_end-1)]     <- list(gpar(fill = "#DDE7F7"))
row_fill[(yellow_start-1):(yellow_end-1)] <- list(gpar(fill = "#FFFACD"))
row_fill[(pink_start-1):(pink_end-1)]     <- list(gpar(fill = "#F8D7DA"))


###########################################
### 4. Build forestplot object (do not immediately open svg)
###########################################
fp <- forestplot(
  labeltext  = tabletext_mod,
  mean       = mean_mod,
  lower      = lower_mod,
  upper      = upper_mod,
  boxsize    = boxsize_mod,
  zero       = 1,
  xlog       = TRUE,
  xticks     = c(0.5, 1, 2, 3, 4),
  xlab       = "Adjusted Odds Ratio (log scale)",
  col        = fpColors(box = "red", line = "darkred", zero = "black"),
  txt_gp     = fpTxtGp(
    label = list(
      gpar(fontsize = 15, fontface = "bold", fontfamily = "serif"),  # header column
      gpar(fontsize = 14, fontface = "bold", fontfamily = "serif"),
      gpar(fontsize = 14, fontface = "bold", fontfamily = "serif"),
      gpar(fontsize = 14, fontface = "bold", fontfamily = "serif")
    ),
    ticks = gpar(fontsize = 30),
    xlab  = gpar(fontsize = 25, fontface = "bold"),
    summary = gpar(fontface = "bold")
  ),
  lineheight = lineheight_vec,
  colgap = unit(5, "mm"),
  align = rep("l", ncol(tabletext_mod)),
  is.summary = c(TRUE, rep(FALSE, length(mean_mod) - 1)),
  ci.vertices = TRUE,
  ci.vertices.height = 0.1,
  new_page = FALSE     # don't force new page yet
)

###########################################
### 5. Apply exposure-group background colors CORRECTLY
###    (pass fp as first argument)
###########################################
# Convert row_fill (list of gpar objects) to argument list and prepend fp
args_list <- c(list(fp), as.list(row_fill))
fp <- do.call(fp_set_zebra_style, args = args_list)

###########################################
### 6. Write to SVG and draw the styled fp
###########################################
svg(output_SVG_file, width = 11, height = 12)
print(fp)   # print method draws the forestplot object with zebra styles
dev.off()
