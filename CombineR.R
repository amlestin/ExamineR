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
    output.file.name.vector <-
      unlist(strsplit(output.file.name, ""))
    period.index <- max(which(output.file.name.vector == "."))
    extensionless.output.file.name <-
      substr(output.file.name, 1, period.index - 1)
    first.and.last.names <-
      paste(unlist(strsplit(extensionless.output.file.name, "-"))[3], collapse = " ")
    output.file.name <-
      paste(first.and.last.names, "Combined Report.xlsx", collapse = "")
    
    column.labels <- c("#",	"Student Response",	"")
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
      current.report <-
        read.csv(report, stringsAsFactors = FALSE, header = FALSE) # read in a report
      current.report <-
        current.report[-1,] # remove the column headers e.g. Question Answer Points.Earned
      
      split.report.name <-  unlist(strsplit(basename(report), "-"))
      course.title <- split.report.name[2]
      student.name <- split.report.name[3]
      
      course.title <- gsub("\\.", " ", course.title)
      student.name <- gsub("\\.", " ", student.name)
      student.name <- gsub("csv", "", student.name)
      
      
      course.and.student.name <- paste(course.title, student.name, sep = " - ")
      
      section.header <-
        c("", course.and.student.name, "") # creates an exam title using the input filename
      
      current.report <-
        rbind(section.header,
              c("", "", ""),
              column.labels,
              current.report) # create a modified report
      if (length(combined.report) == 0) {
        combined.report <-
          rbind(combined.report, current.report)
      } else {
        combined.report <-
          rbind(combined.report, "", "", current.report)
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
              fitToHeight = TRUE)
    
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
