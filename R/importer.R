# Multi-line header excel importer
library(tidyverse)
library(readxl)

#' Import Microsoft Excel file with multiple line header as R tibble
#'
#' @param path A string identifing the location of the Microsoft Excel file
#' @param lines The number of lines used as header in the Microsoft Excel file
#' @param .args Additional arguments to pass to readxl::read_excel function
#' @return A tibble with a proper single line header
#' @examples
#' \dontrun{MLHimport("myExcelFileName.xlsx", lines = 3)}

MLHimport <- function(path, lines, .args = list()) {
  args_default <- c(path = path, .args)
  args_db_call <- c(args_default, skip  = lines - 1)
  args_header  <- c(args_default, n_max = lines - 1)

  # Import the data
  db <- do.call(readxl::read_excel, args_db_call)

  # Import the header
  header <- do.call(readxl::read_excel, args_header) %>%
    t() %>% data.frame() %>% rownames_to_column() %>% mutate_all(as.character)

  # Improve column readability
  names(header) <- paste("line", as.character(1:lines), sep = "_")

  # Turn placeholders to missing
  placeholders <- grepl("^\\.\\.\\.\\d*$", header[, 1])
  header[placeholders, 1] <- NA

  # Fill placeholders NAs
  header <- fill(header, 1)

  # Fill other NAs
  for(line in 1:lines) {
    current_line <- paste("line", as.character(line), sep = "_")
    next_line    <- paste("line", as.character(line + 1), sep = "_")

    if ((line + 1) >= lines) break()

    header <- header %>%
      group_by(.dots = current_line) %>%
      fill(next_line) %>%
      ungroup()
  }

  # Melt columns
  melted_col <- by(header, seq_len(nrow(header)), function(x) {
    paste((x), collapse = "_")})

  col_names <- str_replace_all(melted_col, "_NA", "")

  # Assign new column names
  names(db) <- col_names
  return(db)
}


