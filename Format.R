format.student.csv <- function(input.file) {
#install.packages("xlsx", dependencies = TRUE)
library(xlsx)

student.sheet <- read.csv(input.file)

student.sheet.df <- student.sheet

# creates an xlsx workbook with data from student.sheet
wb <- createWorkbook(type="xlsx")
sheet <- createSheet(wb, sheetName = "Sheet1")
addDataFrame(student.sheet.df, sheet, startRow = 1, startColumn = 1, row.names = FALSE)

# creates formatting objects
red.fill <- Fill(foregroundColor = "red")
zero.cell.style <- CellStyle(wb, fill = red.fill)

# gets data below the title (row 1) and in column 2
sheet.rows <- getRows(sheet = sheet, rowIndex=2:(nrow(df)+1))
sheet.cells <- getCells(sheet.rows, colIndex = 3)
values <- lapply(sheet.cells, getCellValue)


# find the cells that fit the formatting criteria
cells.to.style <- c()
for (i in names(values)) {
  x <- as.numeric(values[i])
  if (x == 0 && !is.na(x)) {
    cells.to.style <- c(cells.to.style, i)
  }    
}

# apply formatting
lapply(names(sheet.cells[cells.to.style]),
       function(ii) setCellStyle(sheet.cells[[ii]], zero.cell.style))

input.file.name.vector <- unlist(strsplit(input.file.name, ""))
period.index <- which(input.file.name.vector == ".")
extensionless.input.file.name <- substr(input.file.name, 1, period.index - 1)
output.file.name <- paste(extensionless.input.file.name, "Formatted", ".xlsx", sep = "")
saveWorkbook(wb, output.file.name)
}
