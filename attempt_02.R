# Load necessary libraries
library(tidyxl)
library(readxl)

# Specify the path to your Excel file
file_path <- "sample_formats3.xlsx"



# Using tidyxl to get cell formats
cell_formats <- xlsx_formats(file_path)

# Print the formats for some example cells using detailed indexing
# Since direct addresses seem unavailable, we must assume cell positions or check manually
# Let's assume you want formatting details for the first cell in your dataset

# Extracting font properties
font_name <- cell_formats$local$font$name[1]
font_size <- cell_formats$local$font$size[1]
font_color <- cell_formats$local$font$color$rgb[1]
font_bold <- cell_formats$local$font$bold[1]

# Extracting fill properties
fill_color <- ifelse(is.na(cell_formats$local$fill$patternFill$fgColor$rgb[1]), 
                     cell_formats$local$fill$patternFill$bgColor$rgb[1],
                     cell_formats$local$fill$patternFill$fgColor$rgb[1])

# Printing out the formatting details
print(paste("Font Name:", font_name))
print(paste("Font Size:", font_size))
print(paste("Font Color:", font_color))
print(paste("Bold:", font_bold))
print(paste("Fill Color:", fill_color))


# Function to extract the fill color more comprehensively
get_fill_color <- function(cell_index) {
  # Check for solid pattern fill first
  fgColor <- cell_formats$local$fill$patternFill$fgColor$rgb[[cell_index]]
  bgColor <- cell_formats$local$fill$patternFill$bgColor$rgb[[cell_index]]
  if (!is.na(fgColor) && fgColor != "") {
    return(fgColor)
  } else if (!is.na(bgColor) && bgColor != "") {
    return(bgColor)
  }
  
  # Check for gradient fill if pattern fill not found
  gradientColor <- cell_formats$local$fill$gradientFill$stop1$color$rgb[[cell_index]]
  if (!is.na(gradientColor) && gradientColor != "") {
    return(gradientColor)
  }
  
  # Default case if no color found
  return("No color found")
}

# Example: Extracting fill color for the first cell
fill_color <- get_fill_color(1)

# Printing out the fill color
print(paste("Fill Color:", fill_color))


# Print the complete fill attribute details for the first few cells to investigate further
if(length(cell_formats$local$fill) > 0) {
  for (i in 1:min(length(cell_formats$local$fill), 3)) { # Adjust to check more cells if necessary
    print(paste("Details for Cell", i, ":"))
    print(cell_formats$local$fill[[i]])
  }
} else {
  print("No fill details available in the dataset.")
}



# Function to extract the fill color, now correctly targeting sub-fields
get_fill_color <- function(cell_index) {
  fgColor <- cell_formats$local$fill$patternFill$fgColor$rgb[[cell_index]][2]  # Accessing the second element
  bgColor <- cell_formats$local$fill$patternFill$bgColor$rgb[[cell_index]][2]  # Accessing the second element
  if (!is.na(fgColor) && fgColor != "") {
    return(fgColor)
  } else if (!is.na(bgColor) && bgColor != "") {
    return(bgColor)
  } else {
    return("No color found")
  }
}

# Example: Extracting fill color for the first cell
fill_color <- get_fill_color(1)

# Printing out the fill color
print(paste("Fill Color:", fill_color))



cell_formats$local$fill$patternFill$bgColor$rgb[[2]]








# Specify the path to your Excel file
file_path <- "sample_formats3.xlsx"



# Using tidyxl to get cell formats
cell_formats <- xlsx_formats(file_path)

# Print the formats for some example cells using detailed indexing
# Since direct addresses seem unavailable, we must assume cell positions or check manually
# Let's assume you want formatting details for the first cell in your dataset

# Extracting font properties
font_name <- cell_formats$local$font$name[1]
font_size <- cell_formats$local$font$size[1]
font_color <- cell_formats$local$font$color$rgb[1]
font_bold <- cell_formats$local$font$bold[1]

# Extracting fill properties
fill_color <- ifelse(is.na(cell_formats$local$fill$patternFill$fgColor$rgb[1]), 
                     cell_formats$local$fill$patternFill$bgColor$rgb[1],
                     cell_formats$local$fill$patternFill$fgColor$rgb[1])

# Printing out the formatting details
print(paste("Font Name:", font_name))
print(paste("Font Size:", font_size))
print(paste("Font Color:", font_color))
print(paste("Bold:", font_bold))
print(paste("Fill Color:", fill_color))

# Function to extract the fill color, enhanced with debugging information
get_fill_color <- function(cell_index) {
  # Attempt to extract foreground color
  fgColor <- ifelse(length(cell_formats$local$fill$patternFill$fgColor$rgb[[cell_index]]) > 0,
                    cell_formats$local$fill$patternFill$fgColor$rgb[[cell_index]][1],  # Accessing first element if available
                    NA)
  
  # Attempt to extract background color
  bgColor <- ifelse(length(cell_formats$local$fill$patternFill$bgColor$rgb[[cell_index]]) > 0,
                    cell_formats$local$fill$patternFill$bgColor$rgb[[cell_index]][1],  # Accessing first element if available
                    NA)
  
  # Print debug information to understand what is being extracted
  print(paste("Debug - Cell:", cell_index, "FG Color:", fgColor, "BG Color:", bgColor))
  
  # Determine which color to return
  if (!is.na(fgColor) && fgColor != "") {
    return(fgColor)
  } else if (!is.na(bgColor) && bgColor != "") {
    return(bgColor)
  } else {
    return("No color found")
  }
}

# Example: Extracting fill colors for the first few cells to verify the function
for (i in 1:5) {
  fill_color <- get_fill_color(i)
  print(paste("Fill Color for cell", i, ":", fill_color))
}
