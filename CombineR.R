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
      
      text.with.student.id <- gsub("./", "", path.to.current.report)
      text.with.student.id <-
        unlist(strsplit(text.with.student.id, "-"))[2]
      
      student.id <- gsub(".csv", "", text.with.student.id)
      
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
  path.to.examiner.folder <- choose.dir()
  # Mac version
  # path.to.examiner.folder <- .rs.api.selectDirectory()
  id.to.reports.map <- FindReportsById(path.to.examiner.folder)
  number.of.students <- length(names(id.to.reports.map))
  
  progress.bar <- winProgressBar(
    title = "CombineR Progress",
    min = 0,
    max = number.of.students,
    width = 300
  )
  
  for (i in 1:number.of.students) {
    report.paths <- id.to.reports.map[[i]]
    
    #report.paths <- paths.to.reports
    # extracts the input filename and modifies it to create an output filename
    output.file.name <- basename(report.paths[1])
    
    # output.file.name.vector <-
    #   unlist(strsplit(output.file.name, ""))
    # period.index <- max(which(output.file.name.vector == "."))
    # extensionless.output.file.name <-
    #   substr(output.file.name, 1, period.index - 1)
    
    extensionless.output.file.name <-
      gsub(".csv", "", output.file.name)
    first.and.last.names <-
      paste(unlist(strsplit(extensionless.output.file.name, "-"))[1], collapse = " ")
    output.file.name <-
      paste(first.and.last.names, "Combined Report.xlsx", collapse = "")
    
    column.labels <- c("Q#", "Student Response",	"")
    
    wb <- createWorkbook("Admin")
    sheet.number <- 1
    addWorksheet(wb, sheet.number) # add modified report to a worksheet
    
    meta.report.paths <- report.paths
    meta.report.paths <-
      gsub("Medical", meta.report.paths, replacement = "")
    meta.report.paths <-
      gsub("Basic", meta.report.paths, replacement = "")
    sorted.indices.paths <-
      sort(meta.report.paths, index.return = TRUE)$ix
    sorted.report.paths <- report.paths[sorted.indices.paths]
    
    report.paths <- sorted.report.paths
    
    combined.report <- c()
    for (report in report.paths) {
      tryCatch({
        current.report <-
          read.csv(
            report,
            stringsAsFactors = FALSE,
            header = FALSE,
            fileEncoding = "UTF-8"
          ) # read in a report
      }, warning = function(w) {
        print("Warning given when reading CSV as UTF-8")
        current.report <-
          read.csv(report,
                   stringsAsFactors = FALSE,
                   header = FALSE) # read in a report
      })
      
      course.title <-
        gsub("\\.", " ", as.character(current.report[4,][1]))
      student.name <- as.character(current.report[3,][1])
      course.and.student.name <-
        paste(course.title, student.name, sep = " - ")
      
      section.header <-
        c("", course.and.student.name, "") # creates an exam title using the input filename
      
      current.report <-
        current.report[-seq(1, 4),] # remove the column headers e.g. Question Answer Points.Earned
      
      current.report <-
        rbind(section.header,
              column.labels,
              current.report) # create a modified report
      if (length(combined.report) == 0) {
        combined.report <-
          rbind(combined.report, current.report)
      } else {
        combined.report <-
          rbind(combined.report, "", current.report)
      }
      
    }
    
    writeData(wb,
              sheet = sheet.number,
              combined.report,
              colNames = FALSE) # add the new worksheet to the workbook
    
    # style.right.align.scores <- createStyle(halign = "center")
    # style.right.align.scores <- createStyle(fontColour = "blue", halign = "right", valign = "center")
    #
    # conditionalFormatting(wb, sheet.number, cols=2, rows=1:nrow(combined.report), type = "contains", rule="Total Score:", style = style.right.align.scores)
    # conditionalFormatting(wb, sheet.number, cols=2, rows=1:nrow(combined.report), type = "contains", rule="Pts missed:", style = style.right.align.scores)
    #
    #
    # resizes column widths to fit contents
    setColWidths(wb, sheet.number, cols = 1:3, widths = "auto")
    # makes sure sheet fits on one printable page
    pageSetup(wb,
              sheet.number,
              fitToWidth = TRUE,
              fitToHeight = FALSE)
    
    saveWorkbook(wb, output.file.name, overwrite = TRUE) # writes a workbook containing all reports inputted
    
    setWinProgressBar(progress.bar, i, title = paste(round(i / number.of.students *
                                                             100, 0),
                                                     "% done", sep = ""))
  }
  close(progress.bar)
}

OpenDir <- function(dir = getwd()) {
  if (.Platform['OS.type'] == "windows") {
    shell.exec(dir)
  } else {
    system("open .")
  }
}

if (winDialog("okcancel", "Select the ExamineR directory created by ExamineR.R") == "OK") {
  CombineReports()
  OpenDir()
}
