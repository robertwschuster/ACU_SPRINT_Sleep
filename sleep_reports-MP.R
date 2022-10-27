#-----------------------------------------------------------------------------------------
#
# Load and analyze sleep tracking data (*.csv files)
# Robert Schuster (ACU SPRINT)
# October 2022
#
#-----------------------------------------------------------------------------------------


# clear environment
rm(list = ls())


# ENTER VARIABLES ########################################################################

# === RECOMMENDED SLEEP PER NIGHT ===
rsn <- 8

# === SLEEP EFFICIENCY REFERENCE ===
nse <- 85

# # === TIMEZONES ===
# # Please enter the timezone as per the 'TZ database name' column found under the following link:
# # https://en.wikipedia.org/wiki/List_of_tz_database_time_zones
# 
# # timezone in which data was downloaded
# tzd <- 'Australia/Brisbane'
# # timezone in which data was collected (e.g., Europe/London)
# tzc <- 'Australia/Brisbane'

# === MENSTRUAL CYCLE DATES === 
C1 <- as.Date(c("2022-06-28", # start of cycle 1
                "2022-07-25")) # end of cycle 1
C1P2 <- as.Date("2022-07-03") # start of phase 2
C1P3 <- as.Date("2022-07-12") # start of phase 3

C2 <- as.Date(c("2022-07-26", # start of cycle 1
                "2022-08-23")) # end of cycle 2
C2P2 <- as.Date("2022-07-31") # start of phase 2
C2P3 <- as.Date("2022-08-10") # start of phase 3

##########################################################################################


# Libraries ------------------------------------------------------------------------------
if ("xlsx" %in% rownames(installed.packages()) == FALSE) {
  install.packages("xlsx")
}

library(xlsx)


# Functions ------------------------------------------------------------------------------
# Convert minutes to hours:minutes
hm <- function(mins) {
  h <- mins %/% 60
  if (round(mins %% 60) == 60) {
    h <- h + 1
    m <- 0
  } else {
    m <- round(mins %% 60)
  }
  return(sprintf("%.01d:%.02d",h,m))
}

# Menstural cycle figure spreadsheets
cyc <- function(df,P2,P3) {
  colnames(df)[c(2:6,8)] <- paste(colnames(df)[c(2:6,8)],'P1')
  
  df <- cbind(df[,1:2,drop = F],replicate(2,df[,2]),
              df[,3,drop = F],replicate(2,df[,3]),
              df[,4,drop = F],replicate(2,df[,4]),
              df[,5,drop = F],replicate(2,df[,5]),
              df[,6,drop = F],replicate(2,df[,6]),
              df[,7,drop = F],
              df[,8,drop = F],replicate(2,df[,8]),
              df[,9:11,drop = F])
  
  p1 <- grep('P1',colnames(df))
  colnames(df)[c(3,6,9,12,15,19)] <- gsub('1','2',colnames(df)[p1])
  colnames(df)[c(4,7,10,13,16,20)] <- gsub('1','3',colnames(df)[p1])
  
  df[which(df$Date >= P2),grep('P1',colnames(df))] <- NA
  df[which(df$Date < P2 | df$Date >= P3),grep('P2',colnames(df))] <- NA
  df[which(df$Date < P3),grep('P3',colnames(df))] <- NA
  
  return(df)
}

# Time in bed
timeInBed <- function(df,CS,CE) {
  df$`End Date` <- as.Date(df$`End Date`,'%d/%m/%Y')
  
  tib <- matrix(0,6,3)
  colnames(tib) <- c('Average','Minimum','Maximum')
  rownames(tib) <- c('Total','Time to fall asleep','Sleep per night','Time awake','Sleep efficiency (quality - %)','WASO')
  
  if (missing(CS) && missing(CE)) {
    rest <- which(df$`Interval Type` == 'REST')
    sleep <- which(df$`Interval Type` == 'SLEEP')
  } else {
    rest <- which(df$`Interval Type` == 'REST' & (df$`End Date`-1) >= CS & (df$`End Date`-1) <= CE)
    sleep <- which(df$`Interval Type` == 'SLEEP' & (df$`End Date`-1) >= CS & (df$`End Date`-1) <= CE)
  }
  funs <- list(mean, min, max)
  # Total time in bed (h:min)
  tib[1,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = df$Duration[rest])
  # Time to fall asleep (h:min)
  if ('Onset Latency' %in% colnames(df)) {
    tib[2,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = df$`Onset Latency`[sleep])
  }
  # Sleep per night (h:min)
  tib[3,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = df$`Sleep Time`[sleep])
  # Time awake / light sleep per night (h:min)
  tib[4,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = df$`Wake Time`[sleep])
  # Sleep efficiency (quality - %)
  tib[5,1:3] <- sapply(funs, function(fun, x) round(fun(x, na.rm = T)), x = df$Efficiency[sleep])
  # WASO (h:min)
  if ('WASO' %in% colnames(df)) {
    tib[6,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = df$WASO[sleep])
  }
  return(tib)
}

# Mean, min, max bed and wake times
bwt <- function(sld,CS,CE) {
  df <- matrix(0,2,3)
  colnames(df) <- c('Average','Minimum','Maximum')
  rownames(df) <- c('Bed time','Wake-up time')
  if (missing(CS) && missing(CE)) {
    r <- 1:nrow(sld)
  } else {
    r <- which(sld$Date >= CS & sld$Date <= CE)
  }
  # Bed time
  df[1,1] <- format(as.POSIXct(Sys.Date() + mean(sld$`Bed time`[r])),'%I:%M %p', tz = 'GMT')
  df[1,2] <- format(as.POSIXct(Sys.Date() + min(sld$`Bed time`[r])),'%I:%M %p', tz = 'GMT')
  df[1,3] <- format(as.POSIXct(Sys.Date() + max(sld$`Bed time`[r])),'%I:%M %p', tz = 'GMT')
  # Wake-up time
  df[2,1] <- format(as.POSIXct(Sys.Date() + mean(sld$`Wake-up time`[r])),'%I:%M %p', tz = 'GMT')
  df[2,2] <- format(as.POSIXct(Sys.Date() + min(sld$`Wake-up time`[r])),'%I:%M %p', tz = 'GMT')
  df[2,3] <- format(as.POSIXct(Sys.Date() + max(sld$`Wake-up time`[r])),'%I:%M %p', tz = 'GMT')
  
  return(df)
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
  subjects[f] <- gsub('[[:digit:]]+', '', sub("([A-Za-z]+_[A-Za-z0-9]+).*", "\\1", basename(flist[f])))
  if (grepl("_",substr(subjects[f],nchar(subjects[f]),nchar(subjects[f])))) {
    subjects[f] <- gsub('[[:punct:]]', '', subjects[f])
  }
  # Alternative
  # if (grepl('[[:digit:]]+',sub("^([^_]*_){1}([^_]*).*", "\\2", basename(flist[f])))) {
  #   subjects[f] <- sub("^([^_]*){1}(_[^_]*).*", "\\1", basename(flist[f]))
  # } else {
  #   subjects[f] <- gsub('[[:digit:]]+', '', sub("^([^_]*_[^_]*).*", "\\1", basename(flist[f])))
  # }
  
  nCol <- max(count.fields(flist[f][1], sep = ",")) # determine number of columns
  data <- read.csv(flist[f][1], header = F, sep = ",", col.names = paste0("V",seq_len(nCol)), fill = T)
  data <- Filter(function(x)!all(is.na(x)), data) # remove columns that are all NAs
  data <- data[!apply(data == "", 1, all),] # remove empty rows
  
  # Statistics
  temp <- data[(which(grepl("Statistics",data[,1]))+1):(which(grepl("Marker/Score List",data[,1]))-1),]
  colnames(temp) <- temp[1,]
  temp <- temp[-c(1:2),]
  
  # Remove summaries from stats
  srows <- which(grepl("Summary",temp[,1]))
  
  # Merge files
  if (f > 1 && grepl(subjects[f],subjects[1:f-1])) {
    stats[[subjects[f]]] <- rbind(stats[[subjects[f]]], temp[-srows,])
  } else {
    stats[[subjects[f]]] <- temp[-srows,]
  }
  stats[[subjects[f]]] <- stats[[subjects[f]]][
    rowSums(stats[[subjects[f]]][,3:ncol(stats[[subjects[f]]])] == 'NaN') != ncol(stats[[subjects[f]]])-2,] # remove rows that are all NaN
  stats[[subjects[f]]] <- stats[[subjects[f]]][
    rowSums(is.na(stats[[subjects[f]]][,3:ncol(stats[[subjects[f]]])])) != ncol(stats[[subjects[f]]])-2,] # remove rows that are all NA
  
  # convert columns to the right classes
  stats[[subjects[f]]][c(2,7:18)] <- sapply(stats[[subjects[f]]][c(2,7:18)],as.numeric)
}
rm(f,nCol,data,temp,srows)

# Sort merged data
if (any(duplicated(subjects))) {
  sdup <- subjects[anyDuplicated(subjects)]
  for (s in 1:length(sdup)) {
    # sort data based on 'Interval Type' and 'Start Date'
    stats[[sdup[s]]] <- stats[[sdup[s]]][with(stats[[sdup[s]]], order(as.factor(`Interval Type`),
                                                                      as.Date(`Start Date`,'%d/%m/%Y'))),]
    # remove duplicate rows based on 'Interval Type', 'Start Date' and 'End Date'
    stats[[sdup[s]]] <- stats[[sdup[s]]][-which(duplicated(stats[[sdup[s]]][,c('Interval Type','Start Date','End Date')])),]
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
    temp <- data.frame(matrix(NA,5,ncol(stats[[s]])))
    colnames(temp) <- colnames(stats[[s]])
    rownames(temp) <- paste(f,c('n','Minimum(n)','Maximum(n)','Average(n)','Std Dev(n-1)'))
    temp[1,c(7:18)] <- colSums(!is.na(stats[[s]][which(stats[[s]]$`Interval Type` == f),c(7:18)]))
    temp[2,c(7:18)] <- apply(stats[[s]][which(stats[[s]]$`Interval Type` == f),c(7:18)],2, 
                             function(x) ifelse(!all(is.na(x)), min(x, na.rm = T), NA))
    temp[3,c(7:18)] <- apply(stats[[s]][which(stats[[s]]$`Interval Type` == f),c(7:18)],2,
                             function(x) ifelse(!all(is.na(x)), max(x, na.rm = T), NA))
    temp[4,c(7:18)] <- sapply(stats[[s]][which(stats[[s]]$`Interval Type` == f),c(7:18)],
                              function(x) ifelse(!all(is.na(x)), mean(x, na.rm = T), NA))
    temp[5,c(7:18)] <- apply(stats[[s]][which(stats[[s]]$`Interval Type` == f),c(7:18)],2,
                             function(x) ifelse(!all(is.na(x)), sd(x, na.rm = T), NA))
    
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
  tib <- timeInBed(stats[[s]])
  tibC1 <- timeInBed(stats[[s]],C1[1],C1[2])
  tibC2 <- timeInBed(stats[[s]],C2[1],C2[2])
  
  tibC1P1 <- timeInBed(stats[[s]],C1[1],C1P2-1)
  tibC1P2 <- timeInBed(stats[[s]],C1P2,C1P3-1)
  tibC1P3 <- timeInBed(stats[[s]],C1P3,C1[2])
  tibC2P1 <- timeInBed(stats[[s]],C2[1],C2P2-1)
  tibC2P2 <- timeInBed(stats[[s]],C2P2,C2P3-1)
  tibC2P3 <- timeInBed(stats[[s]],C2P3,C2[2])
  
  
  # Sleep graphs data ----------------------------------------------------------------------
  # === Sleep duration & efficiency ===
  slf <- as.data.frame(matrix(0,length(ints),11))
  colnames(slf) <- c('Date','Bed time','Wake-up time','Sleep latency','Sleep duration','Time awake','Recommended sleep',
                     'Sleep efficiency','Normal efficiency','BT Avg','WT Avg')
  sld <- as.data.frame(matrix(0,length(ints),9))
  colnames(sld) <- c('Date','Bed time','Wake-up time','Time in bed','Sleep duration','Sleep latency','Sleep efficiency',
                     'WASO','#Wake bouts')
  
  # dates <- unique(as.Date(stats[[s]]$`End Date`, '%d/%m/%Y')[sleep]-1)
  dates <- as.Date(stats[[s]]$`End Date`, '%d/%m/%Y')-1
  sdates <- vector()
  edates <- vector()
  st <- as.POSIXct(paste(stats[[s]]$`Start Date`,stats[[s]]$`Start Time`), '%d/%m/%Y %I:%M:%S %p', tz = 'GMT')
  et <- as.POSIXct(paste(stats[[s]]$`End Date`,stats[[s]]$`End Time`), '%d/%m/%Y %I:%M:%S %p', tz = 'GMT')
  
  for (i in ints) {
    r <- which(stats[[s]]$`Interval Type` == 'SLEEP' & stats[[s]]$`Interval#` == i)
    # b <- which(stats[[s]]$`Interval Type` == 'REST' & stats[[s]]$`Interval#` == i & stats[[s]]$`%Invalid SW` == 0)
    b <- which(stats[[s]]$`Interval Type` == 'REST' & stats[[s]]$`Interval#` == i)
    
    # Date
    # slf[i,1] <- dates[i]
    slf[i,1] <- dates[r]
    sdates[i] <- as.Date(stats[[s]]$`Start Date`[r], '%d/%m/%Y')
    edates[i] <- as.Date(stats[[s]]$`End Date`[r], '%d/%m/%Y')
    # Bed time
    sld[i,2] <- st[b]
    # Wake-up time
    sld[i,3] <- et[b]
    # Sleep latency
    if ('Onset Latency' %in% colnames(stats[[s]])) {
      slf[i,4] <- sum(stats[[s]]$`Onset Latency`[r], na.rm = T)/60 # hours
      sld[i,6] <- sum(stats[[s]]$`Onset Latency`[r], na.rm = T) # minutes
    }
    # Sleep duration
    slf[i,5] <- sum(stats[[s]]$`Sleep Time`[r], na.rm = T)/60 # hours
    sld[i,5] <- sum(stats[[s]]$`Sleep Time`[r], na.rm = T) # minutes
    # Time in bed (minutes)
    sld[i,4] <- sum(stats[[s]]$Duration[b], na.rm = T)
    # Time awake / light sleep (hours)
    slf[i,6] <- sum(stats[[s]]$`Wake Time`[r], na.rm = T)/60
    # Sleep efficiency (%)
    slf[i,8] <- sld[i,7] <- sum(stats[[s]]$Efficiency[r], na.rm = T)
    # WASO (minutes)
    if ('WASO' %in% colnames(stats[[s]])) {
      sld[i,8] <- sum(stats[[s]]$`WASO`[r], na.rm = T)
    }
    # Number of wake bouts (#)
    if ('#Wake Bouts' %in% colnames(stats[[s]])) {
      sld[i,9] <- sum(stats[[s]]$`#Wake Bouts`[r], na.rm = T)
    }
  }
  rm(i,r)
  slf$Date <- sld$Date <- as.Date(slf$Date, .Date(0))
  sdates <- as.Date(sdates, .Date(0))
  edates <- as.Date(edates, .Date(0))
  # Convert bed and wake times to julian (fraction of day)
  sld$`Bed time` <- sapply(c(1:length(sld$`Bed time`)), function (x)
    as.numeric(julian(as.POSIXct(sld$`Bed time`[x], origin = .Date(0), tz = 'GMT'), origin = sdates[x])))
  sld$`Wake-up time` <- sapply(c(1:length(sld$`Wake-up time`)), function(x)
    as.numeric(julian(as.POSIXct(sld$`Wake-up time`[x], origin = .Date(0), tz = 'GMT'), origin = edates[x])))
  
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
  slf[,7] <- rsn
  # Normal efficiency
  slf[,9] <- nse
  # Average bed time
  slf[,10] <- mean(slf$`Bed time`)
  # Average wake-up time
  slf[,11] <- mean(slf$`Wake-up time`)
  
  slf[slf == 0] <- NA
  
  # === Menstrual cycle figures ===
  Cycle <- rep(NA,nrow(sld))
  Cycle[which(sld$Date >= C1[1] & sld$Date <= C1[2])] <- 1
  Cycle[which(sld$Date >= C2[1] & sld$Date <= C2[2])] <- 2
  
  Phase <- rep(NA,nrow(sld))
  Phase[which(sld$Date >= C1[1] & sld$Date < C1P2)] <- 1
  Phase[which(sld$Date >= C1P2 & sld$Date < C1P3)] <- 2
  Phase[which(sld$Date >= C1P3 & sld$Date <= C1[2])] <- 3
  Phase[which(sld$Date >= C2[1] & sld$Date < C2P2)] <- 1
  Phase[which(sld$Date >= C2P2 & sld$Date < C2P3)] <- 2
  Phase[which(sld$Date >= C2P3 & sld$Date <= C2[2])] <- 3
  
  sld <- cbind(sld[,1,drop = F],Cycle,Phase,sld[,2:9,drop = F])
  
  slfC1 <- slf[which(slf$Date >= C1[1] & slf$Date <= C1[2]),]
  slfC1 <- cyc(slfC1,C1P2,C1P3)
  slfC2 <- slf[which(slf$Date >= C2[1] & slf$Date <= C2[2]),]
  slfC2 <- cyc(slfC2,C2P2,C2P3)
  
  # Mean, min, max bed and wake times
  bt <- bwt(sld)
  btC1 <- bwt(sld,C1[1],C1[2])
  btC2 <- bwt(sld,C2[1],C2[2])
  
  btC1P1 <- bwt(sld,C1[1],C1P2-1)
  btC1P2 <- bwt(sld,C1P2,C1P3-1)
  btC1P3 <- bwt(sld,C1P3,C1[2])
  btC2P1 <- bwt(sld,C2[1],C2P2-1)
  btC2P2 <- bwt(sld,C2P2,C2P3-1)
  btC2P3 <- bwt(sld,C2P3,C2[2])
  
  
  # Export data to report template ---------------------------------------------------------
  # Load workbook
  wb <- loadWorkbook(srtemp)

  # Get worksheets
  sheets <- getSheets(wb)

  # Add name and date
  rows  <- getRows(sheets$`Report - Short`)
  cells <- getCells(rows)
  setCellValue(cells$`8.4`, gsub("_"," ",s))
  setCellValue(cells$`9.4`, Sys.Date())

  # Add summary table
  # dex <- cbind(rbind(bt,tib),rbind(btC1,tibC1),rbind(btC2,tibC2))
  dex <- cbind(rbind(bt,tib),
               rbind(btC1,tibC1),rbind(btC2,tibC2),
               rbind(btC1P1,tibC1P1),rbind(btC1P2,tibC1P2),rbind(btC1P3,tibC1P3),
               rbind(btC2P1,tibC2P1),rbind(btC2P2,tibC2P2),rbind(btC2P3,tibC2P3))
  addDataFrame(dex, sheets$Summary, startRow = 3, startColumn = 2, col.names = F, row.names = F)

  # Add data for figures
  addDataFrame(slf, sheets$Figures, startRow = 2, startColumn = 1, col.names = F, row.names = F, showNA = T)
  addDataFrame(slfC1, sheets$`Figures C1`, startRow = 2, startColumn = 1, col.names = F, row.names = F, showNA = T)
  addDataFrame(slfC2, sheets$`Figures C2`, startRow = 2, startColumn = 1, col.names = F, row.names = F, showNA = T)

  # Add data for analysis
  addDataFrame(sld, sheets$Data, startRow = 2, startColumn = 1, col.names = F, row.names = F, showNA = T)

  # Force cells to update
  wb$setForceFormulaRecalculation(T)

  # Save the workbook
  saveWorkbook(wb, paste0(fp,'/',s,'-sleep_report.xlsx'))
}

