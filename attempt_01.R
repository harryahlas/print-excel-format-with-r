"D:\\github\\print-excel-format-with-r\\sample_formats.xlsx"

install.packages(c("tidyxl", "openxlsx"))

library(tidyxl)
library(openxlsx)

# Load the Excel file
file_path <- "sample_formats.xlsx"

# Import data and formatting
data <- xlsx_cells(file_path)
styles <- xlsx_formats(file_path)

# Filter for header cells (assuming headers are in the first row)
header_styles <- data[data$row == 1, ]

# Create a new workbook
wb <- createWorkbook()

# Add a worksheet
addWorksheet(wb, "Sheet1")

# Write data to the worksheet
writeData(wb, "Sheet1", readxl::read_excel(file_path))

# Apply styles to header cells
for (i in seq_along(header_styles$address)) {
  # Get style for current cell
  cell_style <- createStyle(numFmt = styles[header_styles$local_format_id[i], "number_format"],
                            fontName = styles[header_styles$local_format_id[i], "font_name"],
                            fontSize = styles[header_styles$local_format_id[i], "font_size"],
                            fontColour = styles[header_styles$local_format_id[i], "font_color"],
                            fgFill = styles[header_styles$local_format_id[i], "fill_fg_color"],
                            halign = styles[header_styles$local_format_id[i], "halign"],
                            valign = styles[header_styles$local_format_id[i], "valign"],
                            textDecoration = ifelse(styles[header_styles$local_format_id[i], "underline"], "underline", "none"),
                            border = styles[header_styles$local_format_id[i], "border"],
                            borderStyle = styles[header_styles$local_format_id[i], "border_style"],
                            borderColour = styles[header_styles$local_format_id[i], "border_color"])
  # Apply style to cell
  addStyle(wb, "Sheet1", cell_style, rows = header_styles$row[i], cols = header_styles$col[i], gridExpand = TRUE)
}

# Save the workbook
saveWorkbook(wb, "output.xlsx", overwrite = TRUE)
