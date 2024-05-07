"D:\\github\\print-excel-format-with-r\\sample_formats.xlsx"

# install.packages(c("tidyxl", "openxlsx"))
library(tidyxl)
library(openxlsx)

# Load the Excel file
file_path <- "sample_formats.xlsx"


# Get the cell data and formatting from the Excel file
cell_data <- xlsx_cells(file_path)
styles <- xlsx_formats(file_path)

# Read data using readxl
data <- readxl::read_excel(file_path)

# Create a new workbook
wb <- createWorkbook()

# Add a worksheet
addWorksheet(wb, "Sheet1")

# Write data to the worksheet
writeData(wb, "Sheet1", data)

# Filter for header cells (assuming headers are in the first row)
header_styles <- cell_data[cell_data$row == 1, ]

# Apply styles to header cells
for (i in seq_along(header_styles$address)) {
  format_id <- header_styles$local_format_id[i]
  
  # Debugging: print type and value details
  cat("Index:", i, "Format ID:", format_id, "Type:", typeof(format_id), "\n")
  
  # Ensure format_id is numeric and valid
  if (is.numeric(format_id) && !is.na(format_id) && format_id > 0 && format_id <= nrow(styles)) {
    cell_style <- createStyle(
      fontName = ifelse(!is.na(styles$font_name[format_id]), styles$font_name[format_id], "Calibri"), 
      fontSize = ifelse(!is.na(styles$font_size[format_id]), styles$font_size[format_id], 11),
      fontColour = ifelse(!is.na(styles$font_color[format_id]), styles$font_color[format_id], "black"),
      textDecoration = ifelse(!is.na(styles$underline[format_id]) && styles$underline[format_id], "underline", "none"),
      fgFill = ifelse(!is.na(styles$fill_fg_color[format_id]), styles$fill_fg_color[format_id], "white"),
      halign = ifelse(!is.na(styles$halign[format_id]), styles$halign[format_id], "center"),
      valign = ifelse(!is.na(styles$valign[format_id]), styles$valign[format_id], "center"),
      border = ifelse(!is.na(styles$border_style[format_id]), styles$border_style[format_id], "none"),
      borderColour = ifelse(!is.na(styles$border_color[format_id]), styles$border_color[format_id], "black")
    )
    
    # Apply style to the cell
    addStyle(wb, "Sheet1", cell_style, rows = header_styles$row[i], cols = header_styles$col[i], gridExpand = TRUE)
  } else {
    cat("Skipped due to invalid format ID or types at index:", i, "\n")
  }
}

# Save the workbook
saveWorkbook(wb, "output.xlsx", overwrite = TRUE)