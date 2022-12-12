#-----------------------------------------------------------------------------------------
#
# Load and analyze sleep tracking data (*.csv files)
# Robert Schuster (ACU SPRINT)
# October 2022
#
#-----------------------------------------------------------------------------------------


# clear environment
rm(list = ls())


# ENTER VARIABLES ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# === RECOMMENDED SLEEP PER NIGHT ===
rsn <- 8

# === SLEEP EFFICIENCY REFERENCE ===
nse <- 85

# # === TIMEZONES ===
# # Please enter the timezone as per the 'TZ database name' column found under the following link:
# # https://en.wikipedia.org/wiki/List_of_tz_database_time_zones
# 
# # timezone in which data was downloaded
# tzd <- "Australia/Brisbane"
# # timezone in which data was collected (e.g., Europe/London)
# tzc <- "Australia/Brisbane"

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


# Libraries ------------------------------------------------------------------------------
if ("xlsx" %in% rownames(installed.packages()) == FALSE) {
  install.packages("xlsx")
}

library(xlsx)


# Functions ------------------------------------------------------------------------------
# convert minutes to hours:minutes
hm <- function(mins) {
  h <- mins %/% 60
  if (round(mins %% 60) == 60) {
    h <- h + 1
    m <- 0
  } else {
    m <- round(mins %% 60)
  }
  return(sprintf("%.01d:%.02d", h, m))
}


# Load sleep report excel template -------------------------------------------------------
srtemp <- choose.files(caption = "Select the sleep report template excel file")


# Load data ------------------------------------------------------------------------------
# select files to analyze
flist <- choose.files(caption = "Choose sleep tracking files")

subjects <- character(length(flist))
fp <- dirname(flist[1])
stats <- list()
for (f in 1:length(flist)) {
  # remove numbers and punctuation from filenames
  subjects[f] <- gsub("[[:digit:]]+", "", sub("([A-Za-z]+_[A-Za-z0-9]+).*", "\\1", basename(flist[f])))
  if (grepl("_", substr(subjects[f], nchar(subjects[f]), nchar(subjects[f])))) {
    subjects[f] <- gsub("[[:punct:]]", "", subjects[f])
  }
  
  nCol <- max(count.fields(flist[f][1], sep = ",")) # determine number of columns
  data <- read.csv(flist[f][1], header = FALSE, sep = ",", col.names = paste0("V", seq_len(nCol)), fill = TRUE)
  data <- Filter(function(x) !all(is.na(x)), data) # remove columns that are all NAs
  data <- data[!apply(data == "", 1, all), ] # remove empty rows
  
  # Statistics
  temp <- data[(grep("Statistics", data[,1]) + 1):(grep("Marker/Score List", data[, 1]) - 1), ]
  colnames(temp) <- temp[1, ]
  temp <- temp[-c(1:2), ]
  
  # Remove summaries from stats
  srows <- which(grepl("Summary", temp[, 1]))

  # Merge files
  if (f > 1 && grepl(subjects[f], subjects[1:f - 1])) {
    stats[[subjects[f]]] <- rbind(stats[[subjects[f]]], temp[-srows, ])
  } else {
    stats[[subjects[f]]] <- temp[-srows, ]
  }
  # remove rows that are all NaN
  stats[[subjects[f]]] <- stats[[subjects[f]]][
    rowSums(stats[[subjects[f]]][, 3:ncol(stats[[subjects[f]]])] == "NaN") != ncol(stats[[subjects[f]]]) - 2, ]
  # remove rows that are all NA
  stats[[subjects[f]]] <- stats[[subjects[f]]][
    rowSums(is.na(stats[[subjects[f]]][, 3:ncol(stats[[subjects[f]]])])) != ncol(stats[[subjects[f]]]) - 2, ]
  
  # convert columns to the right classes
  stats[[subjects[f]]][c(2, 7:18)] <- sapply(stats[[subjects[f]]][c(2, 7:18)], as.numeric)
}
rm(f, nCol, data, temp, srows)

# Sort merged data
if (any(duplicated(subjects))) {
  sdup <- subjects[anyDuplicated(subjects)]
  for (s in 1:length(sdup)) {
    # sort data based on "Interval Type" and "Start Date"
    stats[[sdup[s]]] <- stats[[sdup[s]]][with(stats[[sdup[s]]], order(as.factor(`Interval Type`),
                                                                      as.Date(`Start Date`, "%d/%m/%Y"))), ]
    # remove duplicate rows based on "Interval Type", "Start Date" and "End Date"
    stats[[sdup[s]]] <- stats[[sdup[s]]][which(!duplicated(stats[[sdup[s]]][, c("Interval Type","Start Date","End Date")])), ]
    # number intervals
    for (f in unique(stats[[sdup[s]]]$`Interval Type`)) {
      r <- which(stats[[sdup[s]]]$`Interval Type` == f)
      stats[[sdup[s]]]$`Interval#`[r] <- c(1:length(r))
    }
  }
  rm(sdup,s,f,r)
}


# Statistics summary
statsSum <- list()
for (s in names(stats)) {
  for (f in unique(stats[[s]]$`Interval Type`)) {
    temp <- data.frame(matrix(NA, 5, ncol(stats[[s]])))
    colnames(temp) <- colnames(stats[[s]])
    rownames(temp) <- paste(f, c("n","Minimum(n)","Maximum(n)","Average(n)","Std Dev(n-1)"))
    temp[1, c(7:18)] <- colSums(!is.na(stats[[s]][which(stats[[s]]$`Interval Type` == f), c(7:18)]))
    temp[2, c(7:18)] <- apply(stats[[s]][which(stats[[s]]$`Interval Type` == f), c(7:18)], 2, 
                             function(x) ifelse(!all(is.na(x)), min(x, na.rm = TRUE), NA))
    temp[3, c(7:18)] <- apply(stats[[s]][which(stats[[s]]$`Interval Type` == f), c(7:18)], 2,
                             function(x) ifelse(!all(is.na(x)), max(x, na.rm = TRUE), NA))
    temp[4, c(7:18)] <- sapply(stats[[s]][which(stats[[s]]$`Interval Type` == f), c(7:18)],
                              function(x) ifelse(!all(is.na(x)), mean(x, na.rm = TRUE), NA))
    temp[5, c(7:18)] <- apply(stats[[s]][which(stats[[s]]$`Interval Type` == f), c(7:18)], 2,
                             function(x) ifelse(!all(is.na(x)), sd(x, na.rm = TRUE), NA))
    
    if (f == unique(stats[[s]]$`Interval Type`)[1]) {
      statsSum[[s]] <- temp
    } else {
      statsSum[[s]] <- rbind(statsSum[[s]], temp)
    }
  }
}
rm(s,f,temp)


# Sleep report data ----------------------------------------------------------------------
for (s in names(stats)) {
  sleep <- which(stats[[s]]$`Interval Type` == "SLEEP")
  ints <- unique(stats[[s]]$`Interval#`[sleep])
  
  # Time in bed
  tib <- matrix(0, 6, 3)
  colnames(tib) <- c("Average","Minimum","Maximum")
  rownames(tib) <- c("Total","Time to fall asleep","Sleep per night","Time awake","Sleep efficiency (quality - %)","WASO")
  # Total time in bed (h:min)
  tib[1, 1] <- hm(statsSum[[s]]["REST Average(n)", "Duration"])
  tib[1, 2] <- hm(statsSum[[s]]["REST Minimum(n)", "Duration"])
  tib[1, 3] <- hm(statsSum[[s]]["REST Maximum(n)", "Duration"])
  
  # Time to fall asleep (h:min)
  if ("Onset Latency" %in% colnames(statsSum[[s]])) {
    tib[2, 1] <- hm(statsSum[[s]]["SLEEP Average(n)", "Onset Latency"])
    tib[2, 2] <- hm(statsSum[[s]]["SLEEP Minimum(n)", "Onset Latency"])
    tib[2, 3] <- hm(statsSum[[s]]["SLEEP Maximum(n)", "Onset Latency"])
  }
  
  # Sleep per night (h:min)
  tib[3, 1] <- hm(statsSum[[s]]["SLEEP Average(n)", "Sleep Time"])
  tib[3, 2] <- hm(statsSum[[s]]["SLEEP Minimum(n)", "Sleep Time"])
  tib[3, 3] <- hm(statsSum[[s]]["SLEEP Maximum(n)", "Sleep Time"])
  
  # Time awake / light sleep per night (h:min)
  tib[4, 1] <- hm(statsSum[[s]]["SLEEP Average(n)", "Wake Time"])
  tib[4, 2] <- hm(statsSum[[s]]["SLEEP Minimum(n)", "Wake Time"])
  tib[4, 3] <- hm(statsSum[[s]]["SLEEP Maximum(n)", "Wake Time"])
  
  # Sleep efficiency (quality - %)
  tib[5, 1] <- round(statsSum[[s]]["SLEEP Average(n)", "Efficiency"])
  tib[5, 2] <- round(statsSum[[s]]["SLEEP Minimum(n)", "Efficiency"])
  tib[5, 3] <- round(statsSum[[s]]["SLEEP Maximum(n)", "Efficiency"])
  
  # WASO (h:min)
  if ("WASO" %in% colnames(statsSum[[s]])) {
    tib[6, 1] <- hm(statsSum[[s]]["SLEEP Average(n)", "WASO"])
    tib[6, 2] <- hm(statsSum[[s]]["SLEEP Minimum(n)", "WASO"])
    tib[6, 3] <- hm(statsSum[[s]]["SLEEP Maximum(n)", "WASO"])
  }
  
  
# Sleep graphs data ----------------------------------------------------------------------
  # Sleep duration & efficiency
  slf <- as.data.frame(matrix(0, length(ints), 11))
  colnames(slf) <- c("Date","Bed time","Wake-up time","Sleep latency","Sleep duration","Time awake","Recommended sleep",
                     "Sleep efficiency","Normal efficiency","BT Avg","WT Avg")
  sld <- as.data.frame(matrix(0, length(ints), 9))
  colnames(sld) <- c("Date","Bed time","Wake-up time","Time in bed","Sleep duration","Sleep latency","Sleep efficiency",
                     "WASO","#Wake bouts")
  
  dates <- as.Date(stats[[s]]$`End Date`, "%d/%m/%Y") - 1
  sdates <- vector()
  edates <- vector()
  st <- as.POSIXct(paste(stats[[s]]$`Start Date`, stats[[s]]$`Start Time`), "%d/%m/%Y %I:%M:%S %p", tz = "GMT")
  et <- as.POSIXct(paste(stats[[s]]$`End Date`, stats[[s]]$`End Time`), "%d/%m/%Y %I:%M:%S %p", tz = "GMT")
    
  for (i in ints) {
    r <- which(stats[[s]]$`Interval Type` == "SLEEP" & stats[[s]]$`Interval#` == i)
    b <- which(stats[[s]]$`Interval Type` == "REST" & stats[[s]]$`Interval#` == i & stats[[s]]$`%Invalid SW` == 0)
    
    # Date
    slf[i, 1] <- dates[r]
    sdates[i] <- as.Date(stats[[s]]$`Start Date`[r], "%d/%m/%Y")
    edates[i] <- as.Date(stats[[s]]$`End Date`[r], "%d/%m/%Y")
    # Bed time
    sld[i, 2] <- st[b]
    # Wake-up time
    sld[i, 3] <- et[b]
    # Sleep latency
    if ("Onset Latency" %in% colnames(stats[[s]])) {
      slf[i, 4] <- sum(stats[[s]]$`Onset Latency`[r], na.rm = TRUE) / 60 # hours
      sld[i, 6] <- sum(stats[[s]]$`Onset Latency`[r], na.rm = TRUE) # minutes
    }
    # Sleep duration
    slf[i, 5] <- sum(stats[[s]]$`Sleep Time`[r], na.rm = TRUE) / 60 # hours
    sld[i, 5] <- sum(stats[[s]]$`Sleep Time`[r], na.rm = TRUE) # minutes
    # Time in bed (minutes)
    sld[i, 4] <- sum(stats[[s]]$Duration[b], na.rm = TRUE)
    # Time awake / light sleep (hours)
    slf[i, 6] <- sum(stats[[s]]$`Wake Time`[r], na.rm = TRUE) / 60
    # Sleep efficiency (%)
    slf[i, 8] <- sld[i, 7] <- sum(stats[[s]]$Efficiency[r], na.rm = TRUE)
    # WASO (minutes)
    if ("WASO" %in% colnames(stats[[s]])) {
      sld[i, 8] <- sum(stats[[s]]$`WASO`[r], na.rm = TRUE)
    }
    # Number of wake bouts (#)
    if ("#Wake Bouts" %in% colnames(stats[[s]])) {
      sld[i, 9] <- sum(stats[[s]]$`#Wake Bouts`[r], na.rm = TRUE)
    }
  }
  rm(i,r)
  slf$Date <- sld$Date <- as.Date(slf$Date, .Date(0))
  sdates <- as.Date(sdates, .Date(0))
  edates <- as.Date(edates, .Date(0))
  # Convert bed and wake times to julian (fraction of day)
  sld$`Bed time` <- sapply(c(1:length(sld$`Bed time`)), function(x)
    as.numeric(julian(as.POSIXct(sld$`Bed time`[x], origin = .Date(0), tz = "GMT"), origin = sdates[x])))
  sld$`Wake-up time` <- sapply(c(1:length(sld$`Wake-up time`)), function(x)
    as.numeric(julian(as.POSIXct(sld$`Wake-up time`[x], origin = .Date(0), tz = "GMT"), origin = edates[x])))
  
  if (any(sld$`Bed time` < 0.5)) {
    i <- which(sld$`Bed time` < 0.5)
    sld$`Bed time`[i] <- sld$`Bed time`[i] + 1
  }
  if (any(sld$`Wake-up time` > 0.5)) {
    i <- which(sld$`Wake-up time` > 0.5)
    sld$`Wake-up time`[i] <- sld$`Wake-up time`[i] - 1
  }
  slf$`Bed time` <- sld$`Bed time`
  slf$`Wake-up time` <- (sld$`Wake-up time` + 1) - slf$`Bed time` # time difference between bed and wake time (for figure)
  # Recommended sleep per night
  slf[, 7] <- rsn
  # Normal efficiency
  slf[, 9] <- nse
  # Average bed time
  slf[, 10] <- mean(slf$`Bed time`)
  # Average wake-up time
  slf[, 11] <- mean(slf$`Wake-up time`)
  
  slf[slf == 0] <- NA
  
  # Mean, min, max bed and wake times
  bt <- matrix(0, 2, 3)
  colnames(bt) <- c("Average","Minimum","Maximum")
  rownames(bt) <- c("Bed time","Wake-up time")
  # Bed time
  bt[1, 1] <- format(as.POSIXct(Sys.Date() + mean(sld$`Bed time`)), "%I:%M %p", tz = "GMT")
  bt[1, 2] <- format(as.POSIXct(Sys.Date() + min(sld$`Bed time`)), "%I:%M %p", tz = "GMT")
  bt[1, 3] <- format(as.POSIXct(Sys.Date() + max(sld$`Bed time`)), "%I:%M %p", tz = "GMT")
  # Wake-up time
  bt[2, 1] <- format(as.POSIXct(Sys.Date() + mean(sld$`Wake-up time`)), "%I:%M %p", tz = "GMT")
  bt[2, 2] <- format(as.POSIXct(Sys.Date() + min(sld$`Wake-up time`)), "%I:%M %p", tz = "GMT")
  bt[2, 3] <- format(as.POSIXct(Sys.Date() + max(sld$`Wake-up time`)), "%I:%M %p", tz = "GMT")
  
  
# Export data to report template ---------------------------------------------------------
  # Load workbook
  wb <- loadWorkbook(srtemp)
  
  # Get worksheets
  sheets <- getSheets(wb)
  
  # Add name and date
  rows  <- getRows(sheets$`Report - Short`)
  cells <- getCells(rows)
  setCellValue(cells$`8.2`, gsub("_", " ", s))
  setCellValue(cells$`9.2`, Sys.Date())
  
  # Add summary table
  dex <- rbind(bt,tib)
  addDataFrame(dex, sheets$Summary, startRow = 2, startColumn = 2, col.names = FALSE, row.names = FALSE)
  
  # Add data for figures
  addDataFrame(slf, sheets$Figures, startRow = 2, startColumn = 1, col.names = FALSE, row.names = FALSE, showNA = TRUE)
  
  # Add data for analysis
  addDataFrame(sld, sheets$Data, startRow = 2, startColumn = 1, col.names = FALSE, row.names = FALSE, showNA = TRUE)
  
  # Force cells to update
  wb$setForceFormulaRecalculation(TRUE)
  
  # Save the workbook
  saveWorkbook(wb, paste0(fp,"/",s,"-sleep_report.xlsx"))
}
