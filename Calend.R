makeXlsCalendar <- function(year
                            , path2zip = "C:/Options/R/Rtools/bin/zip.exe") {
  ## The functions creates and excel file with a simple year calendar by 3 months in a row
  
  ## openxlsx package is used to manipulate excel file
  library(openxlsx)
  
  ## Setting the sequence of days for the year
  days <- seq(as.Date(paste0(year, "-01-01"))
              , as.Date(paste0(year, "-12-31")), by=1)
  
  ## Creating a data frame with days and their parameters
  calendar <- data.frame(cDay = as.numeric(format(days, "%d"))
                         , cMonth = months(days)
                         , cWeekday = weekdays(days, abbreviate = TRUE)
                         , cWeek = format(days, "%V")
                         , stringsAsFactors = FALSE)
  
  ## Creating a list with wide format for each month
  makeCal <- function(mnth) {
    wd <- reshape(calendar[calendar$cMonth == mnth,]
                  , v.names = "cDay"
                  , idvar = "cWeek"
                  , timevar = "cWeekday"
                  , direction = "wide")
    wdc <- wd[,c(3:9)]
    names(wdc) <- gsub("^.*\\.(.*)$", "\\1", names(wdc))
    wdc <- wdc[,sapply(c("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
                       , function(x) which(names(wdc) == x))]
    wdc
  }
  
  calendarList <- lapply(month.name, makeCal)
  
  ## Create excel workbook and worksheet
  WorkBook <- openxlsx::createWorkbook()
  DataSheet <- addWorksheet(WorkBook, sheetName = "Calendar")
  
  ## Define rows and columns for each month
  firstRow <- 2
  startCols <- c(1,9,17)
  endCols <- startCols + 6
  
  firstMon <- c(1,4,7)
  lastMon <- firstMon + 2
  qtrs <- data.frame(firstMon, lastMon)
  
  ## Defining months heights to get appropriate rows to place consequent months 
  getMax <- function(qtrs) {
    max(sapply(calendarList[qtrs[[1]]:qtrs[[2]]], nrow))
  }
  
  maxMonthsHeights <- apply(qtrs, 1, getMax)
  
  mnthStRow <- firstRow
  allMnthStRows <- data.frame(mnthStRow)
  for (i in seq_along(maxMonthsHeights)) {
    mnthStRow <- mnthStRow + maxMonthsHeights[[i]] + 3
    allMnthStRows <- rbind(allMnthStRows, mnthStRow)
  }
  
  addr <- data.frame(rw = c(apply(allMnthStRows, 1, function(rw) rep(rw, 3)))
                     , cm = rep(startCols, 4))
  
  ## Defining columns for merging cells with months' names
  cols2merge <- data.frame(firstCol = rep(startCols, 4)
                           , lastCol = rep(endCols, 4))
  
  ## Set style to make month names and weekdays in bold
  styleB <- createStyle(textDecoration = "bold")
  
  ## Make used columns narrower
  setColWidths(wb = WorkBook
               , sheet = DataSheet
               , cols = c(1:23)
               , widths = 3.7)
  
  ## Placing all the data to excel
  savexls <- function(i) {
    writeData(wb = WorkBook
              , sheet = DataSheet
              , x = data.frame(calendarList[[i]])
              , startRow = addr$rw[[i]]
              , startCol = addr$cm[[i]]
              , rowNames = FALSE
              , headerStyle = styleB)
    writeData(wb = WorkBook
              , sheet = DataSheet
              , x = month.name[[i]]
              , startRow = addr$rw[[i]]-1
              , startCol = addr$cm[[i]]
              , rowNames = FALSE)
    mergeCells(wb = WorkBook
               , sheet = DataSheet
               , cols = cols2merge$firstCol[[i]]:cols2merge$lastCol[[i]]
               , rows = addr$rw[[i]]-1)
    addStyle(wb = WorkBook
             , sheet = DataSheet
             , rows = addr$rw[[i]]-1
             , cols = addr$cm[[i]]
             , style = styleB)
  }
  
  tmp <- sapply(1:12, savexls)
  
  ## Saving the file
  Sys.setenv("R_ZIPCMD" = path2zip)
  openxlsx::saveWorkbook(WorkBook
                         , sprintf("calendar%s.xlsx", year)
                         , overwrite = TRUE)
}
