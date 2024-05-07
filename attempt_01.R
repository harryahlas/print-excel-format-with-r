"D:\\github\\print-excel-format-with-r\\sample_formats.xlsx"

# install.packages(c("tidyxl", "openxlsx"))
library(tidyxl)
library(openxlsx)

# Load the Excel file
file_path <- "sample_formats.xlsx"


# Get the cell data and formatting from the Excel file
cell_data <- xlsx_cells(file_path)
styles <- xlsx_formats(file_path)


cell_data$character_formatted[2]

# Check and print the structure of 'styles'
if (is.null(styles)) {
  cat("Styles object is null.\n")
} else {
  print(str(styles))
  cat("Number of rows in styles: ", nrow(styles), "\n")
}

# Check if styles is empty or not loaded properly
if (is.null(styles) || is.data.frame(styles) && nrow(styles) == 0) {
  cat("Styles data frame is empty or not loaded properly.\n")
} else {
  cat("Styles loaded: ", nrow(styles), "formats found.\n")
}





# Check and print details of 'styles'
if (is.null(styles) || nrow(styles) == 0) {
  cat("Styles data frame is empty or not loaded properly.\n")
} else {
  cat("Styles loaded: ", nrow(styles), "formats found.\n")
}

# Read data using readxl
data <- readxl::read_excel(file_path)

# Create a new workbook
wb <- createWorkbook()

# Add a worksheet
addWorksheet(wb, "Sheet1")

# Write data to the worksheet
writeData(wb, "Sheet1", data)

# Print detailed information if styles are available
if (!is.null(styles) && nrow(styles) > 0) {
  print(head(styles))  # Print first few rows of 'styles' to see what it contains
} else {
  cat("No styles data available to apply.\n")
}

# Ensure there are styles to apply
if (!is.null(styles) && nrow(styles) > 0) {
  # Filter for header cells (assuming headers are in the first row)
  header_styles <- cell_data[cell_data$row == 1, ]
  
  # Apply styles to header cells
  for (i in seq_along(header_styles$address)) {
    format_id <- header_styles$local_format_id[i]
    cat("Index:", i, "Format ID:", format_id, "Type:", typeof(format_id), "\n")
    
    if (!is.na(format_id) && format_id > 0 && format_id <= nrow(styles)) {
      cell_style <- createStyle(
        # Apply all available styles with fallbacks
        # Your previous style application code here
      )
      # Apply style to the cell
      addStyle(wb, "Sheet1", cell_style, rows = header_styles$row[i], cols = header_styles$col[i], gridExpand = TRUE)
    } else {
      cat("Skipped due to invalid format ID or types at index:", i, "\n")
    }
  }
} else {
  cat("No styles data available to apply.\n")
}

# Save the workbook
saveWorkbook(wb, "output.xlsx", overwrite = TRUE)
















library(tidyxl)
library(openxlsx)
library(tidyverse)

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

# Apply styles to header cells if styles are not null and header_styles is not empty
if (!is.null(styles) && nrow(header_styles) > 0) {
  for (i in seq_along(header_styles$address)) {
    format_id <- header_styles$local_format_id[i]
    
    if (!is.na(format_id) && format_id > 0 && format_id <= length(styles$local$font$underline)) {
      cell_style <- createStyle(
        fontName = "Calibri", # Default to Calibri as example
        fontSize = 11,        # Default size
        fontColour = "black", # Default color
        # textDecoration = ifelse(styles$local$font$underline[format_id], "underline", "none"),
        fgFill = "white",     # Default fill
        halign = "center",    # Default alignment
        valign = "center"     # Default vertical alignment
      )
      
      # Apply style to the cell
      addStyle(wb, "Sheet1", cell_style, rows = header_styles$row[i], cols = header_styles$col[i], gridExpand = TRUE)
    }
  }
  
  # Save the workbook
  saveWorkbook(wb, "output.xlsx", overwrite = TRUE)
} else {
  cat("No styles data available to apply, or header styles are empty.\n")
}


