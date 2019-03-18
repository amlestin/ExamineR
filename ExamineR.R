if (!require("openxlsx", character.only = T, quietly = T)) {
  install.packages("openxlsx")
}

library(openxlsx) # library for reading and creating Excel XLSX files
library(compiler)

OpenDir <- function(dir = getwd()) {
  # function to open a file explorer from R
  # https://stackoverflow.com/questions/12135732/how-to-open-working-directory-directly-from-r-console
  if (.Platform['OS.type'] == "windows") {
    shell.exec(dir)
  } else {
    system("open .")
  }
}

# outputs a report folder for the exam at the path exam_filei
ProcessExam <- function(exam_file) {
  exam.title <- exam_file
  
  # tries to open Canvas file as UTF-8 unless there is a warning
  report <- tryCatch({
    read.csv(exam.title, fileEncoding = "UTF-8")
  }, warning = function(w) {
    print("Reading CSV using default encoding")
    read.csv(exam.title)
  })
  
  # moves into exam directory
  working.dir <- trimws(dirname(exam.title))
  setwd(working.dir)
  
  # extracts exams name from exam file path
  exam.name <- basename(exam.title)
  exam.name <- unlist(strsplit(exam.name, "\\."))[1]
  
  # creates report directory and sets it as the working directory
  if (!dir.exists("ExamineR Reports")) {
    dir.create("ExamineR Reports")
  }
  setwd("ExamineR Reports")
  
  # creates exam directory and sets it as the working directory
  if (!dir.exists(basename(exam.name))) {
    dir.create(basename(exam.name))
  }
  setwd(exam.name)
  
  # retrieves the question columns (those that start with X because they are numbers)
  cols <- colnames(report)
  questions.and.answers <- cols[grep("X", cols)]
  
  # extracts questions from answers using the fact that each question is one answer away from the next and vice-versa
  questions <-
    questions.and.answers[seq(1, length(questions.and.answers), 2)]
  answers <-
    questions.and.answers[seq(2, length(questions.and.answers), 2)]
  
  
  # prompts the user so the questions can be sorted by the number in their respective columns or kept as given
  dialog.message <- paste("Is", exam.name, "randomized?")
  is.randomized <- ifelse(winDialog(type = "yesno", dialog.message) == "YES", TRUE, FALSE)
  if (is.randomized == TRUE) {
    sorted.questions <- questions[order(questions)]
    sorted.answers <- answers[order(questions)]
  } else {
    sorted.questions <- questions
    sorted.answers <- answers
  }
  
  # pre-allocate size of stripped questions vector
  stripped.questions <-
    vector(mode = "character", length = length(sorted.questions))
  
  # extracts vector of question strings from column names
  for (question.index in 1:length(sorted.questions)) {
    stripped.question <-
      unlist(strsplit(sorted.questions[question.index], "\\."))[-1]
    
    stripped.question.string <-
      paste(stripped.question, collapse = ' ')
    
    stripped.question.string <-
      unlist(strwrap(
        stripped.question.string,
        width = 20,
        indent = 5,
        simplify = FALSE
      ))
    
    stripped.question.string <-
      paste(stripped.question[-1], collapse = ' ')
    
    stripped.questions[question.index] <- stripped.question.string
  }
  
  number.of.students <- nrow(report)
  reports <- vector(mode = "list", length = number.of.students)
  for (i in 1:number.of.students) {
    student.report <- c()
    student.answers <- c()
    number.of.questions <- length(sorted.questions)
    question.numbers <- c()
    
    for (j in 1:number.of.questions) {
      # finds the student's answer question j in sorted.questions
      answer.index <- which(colnames(report) == sorted.questions[j])
      
      # ensures only missed questions are written to the final report
      grade <- report[i, answer.index + 1]
      if (grade != "0") {
        next
      }
      
      # adds question number to student's incorrect questions vector
      question.numbers <- c(question.numbers, j)
      
      # adds answer to student's incorrect answers vector
      answer <- as.character(report[i, answer.index])
      answer <- trimws(answer)
      student.answers <- c(student.answers, answer)
    }
    
    if (length(student.answers) == 0) {
      question.numbers <- "ALL"
      student.answers <- "CORRECT"
    }
    
    # combine question numbers and answers
    student.report <-
      cbind(question.numbers, student.answers)

    colnames(student.report) <- NULL
    
    student.name <- as.character(report[i, "name"])
    student.name <- trimws(student.name)
    student.score <- as.character(report[i, "score"])
    
    score.string <- paste("Raw Score", student.score, sep = " = ")
    score.row <- c("", score.string)
    
    student.id <-
      as.character(report[i, "sis_id"]) # gets student ID
    course.title <-
      as.character(report[i, "section"]) # extracts course code and name
    split.course.title <-
      unlist(strsplit(course.title, " ")) # creates character vector from course.title
    course.name <-
      split.course.title[-1] # removes course code from course.title
    course.name <-
      paste(course.name, collapse = " ") # course name as character
    
    student.report <-
      rbind(
        c(student.id, ""),
        c(student.name, ""),
        c(course.name, ""),
        c("#", "Student Response"),
        student.report,
        score.row
      )
    
    class.title <-
      as.character(report[i, "section"]) # extracts class code and name
    split.class.title <-
      unlist(strsplit(class.title, " ")) # creates character vector
    class.name <- split.class.title[-1] # removes class code
    class.name <-
      paste(class.name, collapse = " ") # class name as character
    class.name <- make.names(class.name)
    
    reports[[i]] <- student.report
  }
  
  for (i in 1:number.of.students) {
    student.id <- as.character(reports[[i]][1, 1])
    student.name <- as.character(reports[[i]][2, 1])
    output.file.name <-
      paste(student.name, "-", student.id, ".csv", sep = "")
    write.table(
      reports[[i]],
      output.file.name,
      sep = ",",
      row.names = FALSE,
      col.names = FALSE,
      na = ""
    )
  }
  setwd("..")
}
# compiled version of the above
cProcessExam <- cmpfun(ProcessExam)

ExamineR <- function() {
  if (.Platform['OS.type'] == "windows") {
    input.files <- choose.files()
  } else {
    # only supports one file at a time, in the future may loop through files directory with list.dir
    input.files <- file.choose()
  }
  
  # creates a report for each exam file selected by the user
  number.of.exam.files <- length(input.files)
  
  # initializes a progress bar object window
  progress.bar.title <- "ExamineR Progress: "
  progress.bar <- winProgressBar(
    title = progress.bar.title,
    min = 0,
    max = number.of.exam.files,
    width = 300
  )
  
  # processes all exam files selected by the user
  # increments the progress bar after each exam
  for (exam.count in 1:number.of.exam.files) {
    current.exam <- input.files[exam.count]
    print(paste("Processing ", current.exam, "...", sep = ""))
    cProcessExam(current.exam)
    setWinProgressBar(progress.bar,
                      exam.count,
                      title = paste(
                        progress.bar.title,
                        round(exam.count / number.of.exam.files *
                                100, 0),
                        "% done",
                        sep = ""
                      ))
  }
  
  # removes the progress bar object window
  close(progress.bar)
  setwd("..")
}

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

CreateReport <- function(report) {
  report.paths <- report
  
  # extracts the input filename and modifies it to create an output filename
  output.file.name <- basename(report.paths[1])
  extensionless.output.file.name <-
    gsub(".csv", "", output.file.name)
  first.and.last.names <-
    paste(unlist(strsplit(extensionless.output.file.name, "-"))[1], collapse = " ")
  output.file.name <-
    paste(first.and.last.names, "Combined Report.xlsx", collapse = "")
  
  column.labels <- c("Q#", "Student Response")
  
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
    current.report <- tryCatch({
        read.csv(
          report,
          stringsAsFactors = FALSE,
          header = FALSE,
          fileEncoding = "UTF-8"
        ) # read in a report
    }, warning = function(w) {
      print("Reading CSV using default encoding")
        read.csv(report,
                 stringsAsFactors = FALSE,
                 header = FALSE) # read in a report
    })
    
    course.title <- as.character(current.report[3, ][1])
    student.name <- as.character(current.report[2, ][1])
    course.and.student.name <-
      paste(course.title, student.name, sep = " - ")
    
    section.header <-
      c("", course.and.student.name) # creates an exam title using the input filename
    
    current.report <-
      current.report[-seq(1, 4), ] # remove the column headers e.g. Question Answer Points.Earned
    
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
}

# compiled version of the above
cCreateReport <- cmpfun(CreateReport)

CombineReports <- function(path.to.examiner.folder) {
  #  path.to.examiner.folder <- choose.dir()
  
  path.to.examiner.folder <- getwd()
  
  # Mac version
  # path.to.examiner.folder <- .rs.api.selectDirectory()
  id.to.reports.map <- FindReportsById(path.to.examiner.folder)
  number.of.students <- length(names(id.to.reports.map))
  
  progress.bar.title <- "CombineR Progress: "
  progress.bar <- winProgressBar(
    title = progress.bar.title,
    min = 0,
    max = number.of.students,
    width = 300
  )
  
  for (i in 1:number.of.students) {
    cCreateReport(id.to.reports.map[[i]])
    setWinProgressBar(progress.bar,
                      i,
                      title = paste(
                        progress.bar.title,
                        round(i / number.of.students *
                                100, 0),
                        "% done",
                        sep = ""
                      ))
  }
  close(progress.bar)
}


# runs ExamineR to create reports
ExamineR()

# combines the reports created by ExamineR if user clicks OK
#if (winDialog("okcancel", "Select the ExamineR directory created by ExamineR.R") == "OK") {
if ( (length(which(list.files() == "ExamineR Reports")) > 0) == TRUE ) {
  setwd("ExamineR Reports")
  CombineReports()
  OpenDir()
} else {
  print("No ExamineR Reports directory found. Reports will not be combined.")
}

# quit without warning
formals(quit)$save <- formals(q)$save <- "no"
q(
)