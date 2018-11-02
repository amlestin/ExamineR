# outputs a report folder for the exam at the path exam_file
process_exam <- function(exam_file) {
  exam.title <- exam_file
  
  report <- read.csv(exam.title)
  
  
  working.dir <- trimws(dirname(exam.title))
  setwd(working.dir)
  
  exam.name <- basename(exam.title)
  exam.name <- unlist(strsplit(exam.name, "\\."))[1]
  
  if (!dir.exists("ExamineR Reports")) {
    dir.create("ExamineR Reports")
  }
  
  setwd("ExamineR Reports")
  
  if (!dir.exists(basename(exam.name))) {
    dir.create(basename(exam.name))
  }
  
  setwd(exam.name)
  
  
  cols <- colnames(report)
  questions.and.answers <- cols[grep("X", cols)]
  questions <-
    questions.and.answers[seq(1, length(questions.and.answers), 2)]
  answers <-
    questions.and.answers[seq(2, length(questions.and.answers), 2)]
  
  sorted.questions <- questions[order(questions)]
  sorted.answers <- answers[order(questions)]
  
  # pre-allocate size of stripped questions vector
  stripped.questions <- vector(mode = "character", length = length(sorted.questions))
  
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
  
  student.roster <- c()
  
  for (i in 1:nrow(report)) {
    student.report <- c()
    student.answers <- c()
    grades <- c()
    
    for (j in 1:length(sorted.questions)) {
      answer.index <- which(colnames(report) == sorted.questions[j])
      
      grade <- report[i, answer.index + 1]
      grades <- c(grades, grade)
      
      new.answer <- as.character(report[i, answer.index])
      new.answer <- trimws(new.answer)
      
      student.answers <- c(student.answers, new.answer)
    }
    
    student.name <- as.character(report[i, "name"])
    
    student.roster <- c(student.roster, student.name)
    student.score <- as.character(report[i, "score"])
    student.report <-
      cbind(1:length(student.answers), student.answers, grades)
    
    student.report <-
      rbind(student.report, c("Total Score:", "", student.score))
    colnames(student.report) <-
      c("Question", "Answer", "Points Earned")
    
    
    write.table(
      student.report,
      paste(trimws(student.name), ".csv", sep = ""),
      sep = ",",
      row.names = FALSE,
      na = ""
    )
    
  }
  
}

# function to open a file explorer from R
# https://stackoverflow.com/questions/12135732/how-to-open-working-directory-directly-from-r-console

opendir <- function(dir = getwd()) {
  if (.Platform['OS.type'] == "windows") {
    shell.exec(dir)
  } else {
    system("open .")
  }
}


if (.Platform['OS.type'] == "windows") {
  input.files <- choose.files()
} else {
  # only supports one file at a time, in the future may loop through files directory with list.dir
  input.files <- file.choose()
}

# creates a report for each exam file selected by the user
for (exam.count in 1:length(input.files)) {
  current.exam <- input.files[exam.count]
  print(paste("Processing ", current.exam, "...", sep = ""))
  process_exam(current.exam)
  print("Report processed. Check your input directory for the exam folder.")
  opendir()
}


formals(quit)$save <- formals(q)$save <- "no"
q()