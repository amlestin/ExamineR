if (!require("openxlsx", character.only = T, quietly = T)) {
  install.packages("openxlsx")
}

library(openxlsx) # library for reading and creating Excel XLSX files

FindReportsById <- function(path.to.examiner.folder) {
  setwd(path.to.examiner.folder)
  
  dirs.to.check <- list.dirs()[-1]
  
  id.to.reports.map <- list()
  
  for (dir in dirs.to.check) {
    setwd(dir)
    
    files.in.dir <- list.files(full.names = TRUE)
    
    for (report in files.in.dir) {
      path.to.current.report <- report
      
      text.with.student.id <-
        unlist(strsplit(path.to.current.report, "-"))[1]
      student.id <-
        substr(
          text.with.student.id,
          nchar(text.with.student.id) - 9 + 1,
          nchar(text.with.student.id)
        )
      
      
      long.path.to.current.report <-
        paste(dir, substr(
          path.to.current.report,
          2,
          nchar(path.to.current.report)
        ), sep = "")
      
      id.to.reports.map[[student.id]] <-
        c(id.to.reports.map[[student.id]], long.path.to.current.report)
    }
    
    setwd("..")
  }
  
  return(id.to.reports.map)
}

CombineReports <- function() {
  path.to.examiner.folder <- choose.dir() #DELETE ME
  id.to.reports.map <- FindReportsById(path.to.examiner.folder)
  
  for (i in 1:length(names(id.to.reports.map))) {
    report.paths <- id.to.reports.map[[i]]
    
    #report.paths <- paths.to.reports
    # extracts the input filename and modifies it to create an output filename
    output.file.name <- basename(report.paths[1])
    output.file.name.vector <-
      unlist(strsplit(output.file.name, ""))
    period.index <- max(which(output.file.name.vector == "."))
    extensionless.output.file.name <-
      substr(output.file.name, 1, period.index - 1)
    first.and.last.names <-
      paste(unlist(strsplit(extensionless.output.file.name, "-"))[3], collapse = " ")
    output.file.name <-
      paste(first.and.last.names, "Combined Report.xlsx", collapse = "")
    
    column.labels <- c("Question",	"Answer",	"Points Earned")
    wb <- createWorkbook("Admin")
    sheet.number <- 1
    
    meta.report.paths <- report.paths
    meta.report.paths <-
      gsub("Medical", meta.report.paths, replacement = "")
    meta.report.paths <-
      gsub("Basic", meta.report.paths, replacement = "")
    sorted.indices.paths <-
      sort(meta.report.paths, index.return = TRUE)$ix
    sorted.report.paths <- report.paths[sorted.indices.paths]
    
    report.paths <- sorted.report.paths
    
    for (report in report.paths) {
      current.report <-
        read.csv(report, stringsAsFactors = FALSE, header = FALSE) # read in a report
      current.report <-
        current.report[-1, ] # remove the column headers e.g. Question Answer Points.Earned
      
      section.header <-
        c("", basename(report), "") # creates an exam title using the input filename
      
      current.report <-
        rbind(section.header,
              c("", "", ""),
              column.labels,
              current.report) # create a modified report
      
      addWorksheet(wb, sheet.number) # add modified report to a worksheet
      
      # formats the questions with zero points to be more visible
      zero.points.style <- createStyle(bgFill = "#FFC7CE")
      conditionalFormatting(
        wb,
        sheet.number,
        cols = 3,
        rows = (4:nrow(current.report)),
        type = "contains",
        rule = "0",
        style = zero.points.style
      )
      
      writeData(wb,
                sheet = sheet.number,
                current.report,
                colNames = FALSE) # add the new worksheet to the workbook
      
      # resizes column widths to fit contents
      setColWidths(wb, sheet.number, cols = 1:3, widths = "auto")
      # makes sure sheet fits on one printable page
      pageSetup(wb,
                sheet.number,
                fitToWidth = TRUE,
                fitToHeight = TRUE)
      
      sheet.number <-
        sheet.number + 1 # increment sheet number for next report
    }
    
    
    saveWorkbook(wb, output.file.name, overwrite = TRUE) # writes a workbook containing all reports inputted
  }
}

if (winDialog("okcancel", "Select the ExamineR directory created by ExamineR.R") == "OK") {
  CombineReports()
  winDialog("ok", "Your combined reports are in the ExamineR folder")
}
