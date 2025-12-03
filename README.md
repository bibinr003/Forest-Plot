Explaining the R Forest Plot Script for even a beginner or a non-programmer (Line-by-line)

The input data was from an AQI and Pollutant impact research related to a heart development study from a hospital
#the example data is different from the true study outcomes 

Purpose: This document explains every section and important line of the R script you provided. It is written for someone who does not program, using simple language and short examples. At the end, there are tips on how to change common settings (e.g., file paths, colors, and groups).
1. Top: Libraries — “tools the script needs”
Code lines:

library(readxl)
library(forestplot)
library(dplyr)
library(stringr)
library(grid)

What this means in simple terms:
- The script tells R to load helper packages (like using apps on your phone). Each package gives the script abilities:
  • readxl: open Excel files.
  • forestplot: draw forest plots (those tidy OR-with-CI plots).
  • dplyr: handle and transform table data (sorting, selecting columns).
  • stringr: easy text (string) operations.
  • grid: fine control over drawing shapes and spacing.

Example: If you want to read an Excel file, readxl is the tool that knows how to open it.
2. User-editable block — change these for your file
Code lines (example):

file_path  <- "C:/Users/Bibin/Downloads/Forestplot_AQI_study_22_november.xlsx"
#sheet_name <- "complexXcontrols"
sheet_name <- "acyanotic"
output_SVG_file <- paste0("C:/Users/Bibin/Downloads/Mahima/Nov22_", sheet_name, ".svg")

- This section is where you tell the script which Excel file and which Excel sheet to use. Think of it like pointing to the door the data is behind.
- file_path: full location of the Excel file on your computer.
- sheet_name: which sheet (tab) inside that Excel file.
- output_SVG_file: where to save the plot image; paste0(...) simply joins pieces (folder + sheet name + .svg).

Example: If your Excel is on your Desktop named study.xlsx and the sheet is "results", set file_path to "C:/Users/You/Desktop/study.xlsx" and sheet_name to "results".
3. Read Excel and clean column names
Code lines:

df <- read_excel(file_path, sheet = sheet_name)
colnames(df) <- trimws(colnames(df))

- read_excel(...) opens the chosen sheet and puts it into a table called df (short for data frame).
- trimws(...) removes stray spaces at the start or end of column names (so the script can find them reliably).

Example: If a column header accidentally has a space at the end like "POLLUTANTS ", trimws makes it "POLLUTANTS".
4. Keep only rows that have an Adjusted OR
Code lines:

df <- df %>% filter(`ADJUSTED OR` != "_")

- This removes rows where the ADJUSTED OR column has an underscore "_" (used to mark missing/empty values).

Example: If a row has "_", meaning there was no OR calculated for that exposure, it gets removed before plotting.
5. Extract numeric OR and 95% CI from the ADJUSTED OR text
Code lines:

df <- df %>%
  mutate(
    OR_val    = as.numeric(str_extract(`ADJUSTED OR`, "^[0-9.]+")),
    lower_CI  = as.numeric(str_extract(`ADJUSTED OR`, "(?<=\\().+?(?=-)")),
    upper_CI  = as.numeric(str_extract(`ADJUSTED OR`, "(?<=-).+?(?=\\))"))
  )

- The ADJUSTED OR column likely looks like: 1.35 (0.98-1.86). The script pulls three numbers out:
  • OR_val = 1.35 (the central estimate),
  • lower_CI = 0.98, and
  • upper_CI = 1.86.
- It uses tiny patterns to find numbers before parentheses and inside parentheses separated by a dash.

Simple example: From "2.00 (1.20-3.30)" it extracts OR_val=2.00, lower_CI=1.20, upper_CI=3.30.
6. Extract 4 values from case/control pairs
Code lines:

df <- df %>%
  mutate(
    bad_a  = as.numeric(sub("/.", "", `CASE/CONTROLS(Bad outcome)`)),
    bad_b  = as.numeric(sub(".*/", "", `CASE/CONTROLS(Bad outcome)`)),
    good_c = as.numeric(sub("/.*/,", "", `CASE/CONTROLS(Good outcome)`)),
    good_d = as.numeric(sub(".*/", "", `CASE/CONTROLS(Good outcome)`)),
    total_n = bad_a + bad_b + good_c + good_d
  )

Correction note:
- The script takes two columns: CASE/CONTROLS(Bad outcome) and CASE/CONTROLS(Good outcome). These columns probably look like "5/45" meaning 5 cases and 45 controls.
- The script splits those strings into two numbers each: bad_a and bad_b, good_c and good_d. Then it sums them to get total_n (the row's total sample size).

Corrected simple example: If CASE/CONTROLS(Bad outcome) has "5/45" then bad_a=5 and bad_b=45. If CASE/CONTROLS(Good outcome) has "10/80" then good_c=10 and good_d=80. total_n = 5+45+10+80 = 140.
7. Compute box size based on total sample size
Code lines:

max_boxsize <- 0.8
df <- df %>%
  mutate(
    boxsize_weight = total_n / max(total_n, na.rm = TRUE) * max_boxsize
  )

- The forest plot can show box sizes for each row: larger boxes for rows with more people (more reliable estimates).
- max_boxsize sets the biggest box size to 0.8. The formula scales each row's box relative to the largest total_n in the table.

Example: If the largest row has total_n = 200, a row with total_n = 100 gets boxsize_weight = 100/200 * 0.8 = 0.4.
8. Prepare labels (text shown in the table part of the plot)
Code lines:

OR_CI <- sprintf("%.2f (%.2f–%.2f)", df$OR_val, df$lower_CI, df$upper_CI)

- Creates a neat text like "1.35 (0.98–1.86)" with two decimal places for display in the left table next to the plot.

Example: OR_val=1.23456 becomes "1.23" in the label.
9. Build tabletext (columns that appear left of the plot)
Code lines:

tabletext <- cbind(
  c("Exposure Window", df$`EXPOSURE PERIOD`),
  c("Pollutant", df$POLLUTANTS),
  c("Exposure Level", df$`EXPOSURE LEVEL`),
  c("Adjusted OR (95% CI)", OR_CI)
)

- cbind(...) builds a matrix where each column is one of the columns you want to show next to the forest plot. The first row contains column headers (the labels shown at the top).

Example: The left table will have four columns: Exposure Window, Pollutant, Exposure Level, and Adjusted OR (95% CI).
10. Insert blank rows to group PRE/CONCEPTION/POST
Code lines (concept):

insert_after <- which(df$`EXPOSURE PERIOD` %in% c("PRECONCEPTION", "CONCEPTION"))

- The script finds positions where PRECONCEPTION or CONCEPTION occur and inserts a blank row after them so the plot looks grouped.

Example: If the sheet lists PRECONCEPTION items first, then CONCEPTION, a blank row will be added after each of those groups to visually separate groups.
11. Loop to build the modified table and numeric vectors
Code lines (concept):

A loop goes over each row of the tabletext and builds new objects:
- tabletext_mod: the table text with inserted blanks
- mean_mod, lower_mod, upper_mod: numeric vectors matching table rows
- boxsize_mod: box sizes to use for each row

- This creates the final inputs needed by the forestplot function. The numeric vectors must line up with the text rows (including blanks).
12. Remove accidental empty rows and add one intentional blank
Code lines (concept):

The script removes rows that are completely empty (both text and numbers) so the plot does not have accidental gaps.
Then it adds one extra blank row at the end for bottom padding (aesthetic spacing).

- Think of trimming white space and then adding one tidy blank line at the bottom so the figure does not look cramped.
13. Line heights and bottom padding
Code lines:

bottom_padding <- 0.8
lineheight_vec <- rep(unit(1, "cm"), length(mean_mod))
lineheight_vec[length(lineheight_vec)] <- unit(bottom_padding, "cm")

- This controls the vertical spacing for each row. Each row gets 1 cm unless the last row gets smaller padding controlled by bottom_padding.
14. Row-based coloring (zebra striping and highlights)
Code lines (concept):

n_rows <- nrow(tabletext_mod)
row_fill <- rep(list(gpar(fill = "white")), n_rows)

# user chooses ranges like blue_start <- 2; blue_end <- 5
# then those ranges are painted with the chosen colors

- The script builds a list of colors for each row. By setting start and end numbers you color groups of rows blue, yellow, pink, etc. This is purely visual—useful to highlight exposure groups.

Example: If you want rows 2 to 5 in pale blue, set blue_start = 2 and blue_end = 5.
15. Create the forest plot object (the core drawing command)
Code lines (important arguments):

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
  txt_gp     = fpTxtGp(...),
  lineheight = lineheight_vec,
  colgap = unit(5, "mm"),
  align = rep("l", ncol(tabletext_mod)),
  is.summary = c(TRUE, rep(FALSE, length(mean_mod) - 1)),
  ci.vertices = TRUE,
  ci.vertices.height = 0.1,
  new_page = FALSE
)

Key arguments:
- labeltext: the table on the left.
- mean/lower/upper: numbers for plotting the central estimate and its CI.
- boxsize: size of the square for each estimate.
- zero: the reference value (1 for Odds Ratio).
- xlog: use a logarithmic x scale (ORs are plotted on log scale by convention).
- xticks: tick marks on the x-axis (numbers to show).
- col: colors for boxes and lines.
- txt_gp: font sizes for labels and axis. (The script sets quite large fonts.)
- is.summary: tells the plotting function 'this first row is a header/summary, show it differently'.

Simple example: If a row has mean=2 and CI (1.2,3.3), it draws a square centered at 2 and a horizontal line from 1.2 to 3.3.
16. Apply the row colors and finalize
Code lines:

args_list <- c(list(fp), as.list(row_fill))
fp <- do.call(fp_set_zebra_style, args = args_list)

svg(output_SVG_file, width = 11, height = 12)
print(fp)
dev.off()

- fp_set_zebra_style uses the colors prepared earlier to paint each row.
- svg(...) opens a device (a file) to save the drawing as an SVG image.
- print(fp) actually draws the figure into that file.
- dev.off() closes the file so it finishes writing.

Example: After running, open the SVG file with a browser to view the high-resolution image.
17. Common pitfalls and quick fixes
- Missing packages: If R reports 'package not found', install the package in R once using install.packages("packageName") or Bioconductor instructions for special packages.
- Column names mismatch: Make sure the Excel column headers exactly match names used in the script: ADJUSTED OR, EXPOSURE PERIOD, POLLUTANTS, EXPOSURE LEVEL, CASE/CONTROLS(Bad outcome), CASE/CONTROLS(Good outcome).
- Different formats: If ADJUSTED OR uses a different format (e.g., uses commas or uses 'to' between CI bounds), the regular expressions (text extract rules) must be adapted.
- Space or extra characters: trimws helps for spaces, but odd characters will break numeric conversion.

18. How to change colors, fonts, and axis labels (quick)
- Change box and line colors: modify fpColors(box = "red", line = "darkred"). Use any color name or hex code like "#FF0000".
- Change x-axis ticks: xticks = c(0.5, 1, 2, 4) — edit the vector to values you prefer.
- Change font sizes: inside txt_gp (e.g., gpar(fontsize = 14)). Reduce numbers to make text smaller.
- Change output file type: instead of svg(...), use png("file.png", width=1200, height=1200) to save raster images.

19. Short annotated script (compact)
This is a concise summary of what each main block does:
 1) Load helper packages.
 2) Set file path and sheet name.
 3) Read Excel input and clean headers.
 4) Remove rows with no OR.
 5) Extract numeric OR and CI from text.
 6) Extract case/control numbers and compute total N.
 7) Compute box sizes proportional to N.
 8) Build left-side table text (tabletext).
 9) Insert visual group breaks and create modified vectors.
 10) Configure colors and spacing.
 11) Build the forestplot object and save to file.

20. If something goes wrong — checklist
1. Are package names spelled correctly? Install if missing.
2. Does the Excel file path exist? Copy-paste the full path.
3. Do the Excel columns match the expected names exactly? (Case-sensitive).
4. If the ADJUSTED OR format differs, show one example row and ask for help to update the extraction patterns.

