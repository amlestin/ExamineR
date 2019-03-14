# outputs a report folder for the exam at the path exam_file
ProcessExam <- function(exam_file) {
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
  stripped.questions <-
    vector(mode = "character", length = length(sorted.questions))
  
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
    
    question.numbers <- c()
    
    for (j in 1:length(sorted.questions)) {
      answer.index <- which(colnames(report) == sorted.questions[j])
      grade <- report[i, answer.index + 1]
      
      if (grade != "0") {
        next
      }
      question.numbers <- c(question.numbers, j)
      
      grades <- c(grades, grade)
      
      new.answer <- as.character(report[i, answer.index])
      new.answer <- trimws(new.answer)
      
      student.answers <- c(student.answers, new.answer)
    }
    
    student.name <- as.character(report[i, "name"])
    student.roster <- c(student.roster, student.name)
    student.score <- as.character(report[i, "score"])
    score.points <- paste("Raw Score", student.score, sep = " = ")
    points.missed <-
      paste ("", as.character(report[i, ""]), sep = " ")
    
    student.report <-
      cbind(question.numbers, student.answers, rep(NA, length(question.numbers)))
    
    score.row <- c("", score.points, "")
    missed.pts.row <-
      c("", points.missed, "")
    
    colnames(student.report) <-
      c("#", "Student Response", "")
    
    student.id <-
      as.character(report[i, "sis_id"]) # gets student ID
    
    class.title <-
      as.character(report[i, "section"]) # extracts class code and name
    split.class.title <-
      unlist(strsplit(class.title, " ")) # creates character vector
    class.name <- split.class.title[-1] # removes class code
    class.name <-
      paste(class.name, collapse = " ") # class name as character
    class.name <- make.names(class.name)
    
    student.report <-
      rbind(c( student.id, "",""), c(class.name, "",""), c(trimws(student.name), "","") , student.report, score.row, missed.pts.row)
    
    write.table(
      student.report,
      # paste(
      #   student.id,
      #   "-",
      #   class.name,
      #   "-",
      #   trimws(student.name),
      #   ".csv",
      #   sep = ""
      # ),
      paste(student.id,
            ".csv",
            sep = ""),
      sep = ",",
      row.names = FALSE,
      na = ""
    )
  }
  setwd("..")
}

OpenDir <- function(dir = getwd()) {
  # function to open a file explorer from R
  # https://stackoverflow.com/questions/12135732/how-to-open-working-directory-directly-from-r-console
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
number.of.exam.files <- length(input.files)

# initializes a progress bar object window
progress.bar <- winProgressBar(
  title = "ExamineR Progress",
  min = 0,
  max = number.of.exam.files,
  width = 300
)

# processes all exam files selected by the user
# increments the progress bar after each exam
for (exam.count in 1:number.of.exam.files) {
  current.exam <- input.files[exam.count]
  print(paste("Processing ", current.exam, "...", sep = ""))
  ProcessExam(current.exam)
  setWinProgressBar(progress.bar, exam.count, title = paste(
    round(exam.count / number.of.exam.files *
            100, 0),
    "% done",
    sep = ""
  ))
}

# removes the progress bar object window
close(progress.bar)
OpenDir()

# quit without warning
formals(quit)$save <- formals(q)$save <- "no"
q()