#-----------------------------------------------------------------------------------------
#
# Load and analyze somfit sleep tracking data (*.rtf files)
# Robert Schuster (ACU SPRINT)
# November 2022
#
#-----------------------------------------------------------------------------------------


# clear environment
rm(list = ls())


# Libraries ------------------------------------------------------------------------------
if ("striprtf" %in% rownames(installed.packages()) == F) {
  install.packages("striprtf")
}

if ("xlsx" %in% rownames(installed.packages()) == FALSE) {
  install.packages("xlsx")
}

if ("tidyverse" %in% rownames(installed.packages()) == FALSE) {
  install.packages("tidyverse")
}

library(tidyverse)
library(striprtf)
library(xlsx)


# Basic functions ------------------------------------------------------------------------
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

# import and clean up RTF file
rtf2df <- function(file) {
  df <- read_rtf(file, row_start = "|", row_end = "", cell_end = "|")
  ug <- grep('\u00d9',df)
  # if kiloohm ('\u2126) was replaced with U-grave ('\u00d9') assume that en dash ('\u2013) was not split incorrectly,
  # otherwise assume the opposite
  if (!identical(ug,integer(0))) {
    df[ug-1] <- paste0(df[ug-1],'\u2126||')
    df <- df[-c(ug,ug+1)]
  } else {
    ed <- grep('\u2013',df)
    if (any(diff(ed <= 2))) {
      minus <- which(diff(ed) <= 2)
      rp <- gsub('\\|','',paste0(df[ed[minus]],df[ed[minus]+1],df[ed[minus+1]],df[ed[minus+1]+1]))
      df[ed[minus]-1] <- paste0(df[ed[minus]-1],rp,'|')
      ed <- ed[-c(minus,minus+1)]
    }
    rp <- paste0(sub('\\|','',df[ed]),sub('\\|','',df[ed+1]))
    df[ed-1] <- paste0(df[ed-1],rp)
    df <- df[-c(ed,ed+1)]
  }
  df <- lapply(df, function(x) strsplit(x, split = '\\|')[[1]])
  df <- Filter(length, df)
  df <- lapply(df, function(x) x[2:length(x)]) # remove empty elements
  # convert to dataframe
  df <- data.frame(matrix(unlist(df), nrow = length(df), byrow = T), stringsAsFactors = F)
  colnames(df) <- df[1,]
  df <- df[-1,]
  # find header rows
  df[df == " "] <- "" # change cells with space to empty cells
  hd <- which(apply(df[,2:4],1,function(x) all(x == "")))
  # split dataframe by header rows
  df <- split(df, cumsum(1:nrow(df) %in% as.numeric(hd)))
  names(df) <- lapply(df, '[[', 1,1)
  df <- lapply(df, function(x) x[-1,])
}

# retrieve and convert values
somfind <- function(data, value, table, proc = NULL) {
  rv <- data[[table]]$`Metric Value`[data[[table]]$`Metric Name` == value]
  if (rv == '-' || rv == '- ' || identical(rv,integer(0))) {
    rv <- NA
  } else {
    if (!is.null(proc)) {
      rv <- rv %>% proc
    }
  }
  return(rv)
}

# Sleep metrics Madi
sleepMetrics_MP <- function(data) {
  df <- list()
  # start date
  df$`Start date` <- sDate <- somfind(data,'Study Date','Referrer',function(x) as.Date(x,format = '%d/%m/%Y'))
  # EEG impedence (right and left)
  df$`EEG impedance (right)` <- somfind(data,'EEG Impedance \u2013 Right Rating','Study Quality')
  df$`EEG impedance (left)` <- somfind(data,'EEG Impedance \u2013 Left Rating','Study Quality')
  # Tttal data received - value (%) and rating
  df$`Total data received (%)` <- somfind(data,'Total Data Received','Study Quality',as.numeric)
  df$`Total data received (rating)`  <- somfind(data,'Total Data Received Rating','Study Quality')
  # total scorable EEG - value (%) and rating
  df$`Total scorable EEG (%)` <- somfind(data,'Total Scorable EEG','Study Quality',as.numeric)
  df$`Total scorable EEG (rating)` <- somfind(data,'Total Scorable EEG Rating','Study Quality')
  # Somfit battery - starting level
  df$`Somfit battery start level (%)` <- somfind(data,'Somfit Start Battery Level ','Study Quality',as.numeric)
  # phone start and end battery level (%)
  df$`Phone battery start level (%)` <- somfind(data,'Phone Start Battery Level','Device',as.numeric)
  df$`Phone battery end level (%)` <- somfind(data,'Phone End Battery Level','Device',as.numeric)
  # Somfit device
  df$`Somfit device` <- somfind(data,'Somfit Device','Device')
  # app start and end date/time
  df$`App start date/time` <- somfind(data,'App Start Date/Time','Device',function(x) as.POSIXct(x, format = '%Y/%m/%d %H-%M-%S', tz = 'GMT'))
  df$`App end date/time` <- somfind(data,'App End Date/Time ','Device',function(x) as.POSIXct(x, format = '%Y/%m/%d %H:%M:%S', tz = 'GMT'))
  # start recording
  sr <- somfind(data,'Start Recording','Referrer',function(x) as.POSIXct(x, format = '%H:%M:%S', tz = 'GMT'))
  if (as.numeric(julian(sr)) - as.numeric(Sys.Date()) < 0.5) {
    df$`Start date` <- sDate <- sDate - 1
  }
  df$`Start recording` <- sr <- as.POSIXct(sub("\\S+", sDate, sr), tz = 'GMT')
  # end recording
  er <- somfind(data,'Stop Recording','Referrer',function(x) as.POSIXct(x, format = '%H:%M:%S'))
  df$`End recording` <- er <- as.POSIXct(sub("\\S+", sDate+1, er), tz = 'GMT')
  # lights out/on
  df$`Lights out` <- sr + somfind(data,'Lights Out Time','Sleep',as.numeric)
  df$`Lights on` <- sr + somfind(data,'Lights On Time','Sleep',as.numeric)
  # sleep onset/offset
  df$`Sleep onset` <- sr + somfind(data,'Sleep Onset','Sleep',as.numeric)
  df$`Sleep offset` <- sr + somfind(data,'Sleep Offset','Sleep',as.numeric)
  # total sleep time
  df$`Total sleep time (min)` <- somfind(data,'Total sleep time (TST)','Sleep',function(x) as.numeric(x) %>% `/`(60))
  # sleep latency
  df$`Sleep latency (min)` <- somfind(data,'Sleep Latency','Sleep',function(x) as.numeric(x) %>% `/`(60))
  # WASO
  df$`WASO (min)` <- somfind(data,'Wake after sleep onset (WASO)','Sleep',function(x) as.numeric(x) %>% `/`(60))
  # sleep efficiency
  df$`Sleep efficiency (%)` <- somfind(data,'Sleep efficiency','Sleep',as.numeric)
  # sleep availability
  df$`Sleep availability time (min)` <- somfind(data,'Total sleep period','Sleep',function(x) as.numeric(x) %>% `/`(60))
  df$`Time available for sleep (min)` <- somfind(data,'Time available for sleep','Sleep',function(x) as.numeric(x) %>% `/`(60))
  df$`Total recording time (min)` <- somfind(data,'Total Recording Time (TRT)','Referrer',function(x) as.numeric(x) %>% `/`(60))
  # sleep stages
  # time (mins)
  stages <- c('Time Awake during sleep Period','N1 Sleep Time','N2 Sleep Time','N1 Sleep Time','N3 Sleep Time',
              'REM Sleep Time','Unsure Time')
  cnames <- c('Wake (min)','N1 (min)','N2 (min)','N1/N2 (min)','N3 (min)','REM (min)','Unscored (min)')
  for (n in 1:length(stages)) {
    df[[cnames[n]]] <- somfind(data,stages[n],'Sleep',function(x) as.numeric(x) %>% `/`(60))
  }
  df$`N1/N2 (min)` <- df$`N1/N2 (min)` + df$`N2 (min)`
  # percentage
  df$`Wake (%)` <- somfind(data,'Total Recording Time (TRT)','Referrer',as.numeric)/df$`Time available for sleep (min)`*100
  stages <- c('Stage 1 / N1 %','Stage 2 / N2 %','Stage 1 / N1 %','Stage 3 / N3 %','REM sleep %')
  cnames <- c('N1 (%)','N2 (%)','N1/N2 (%)','N3 (%)','REM (%)')
  for (n in 1:length(stages)) {
    df[[cnames[n]]] <- somfind(data,stages[n],'Sleep',as.numeric)
  }
  df$`N1/N2 (%)` <- df$`N1/N2 (%)` + df$`N2 (%)`
  df$`Unscored (%)` <- df$`Unscored (min)`/df$`Time available for sleep (min)`*100
  # awakenings
  df$`Awakenings (total)` <- somfind(data,'Number of awakenings','Sleep',as.numeric)
  df$`Awakenings (per hour)` <- somfind(data,'Awakenings Index','Sleep',as.numeric)
  # pAHI
  df$pAHI <- somfind(data,'pAHI Total','Respiratory (PAT / Snore)',as.numeric)
  # SpO2
  df$`SpO2 (awake)` <- somfind(data,'SpO2 awake average','SpO2',as.numeric)
  df$`SpO2 (sleep)` <- somfind(data,'Average SpO2 (Sleep)','SpO2',as.numeric)
  # HRV and pulse
  df$HRV <- somfind(data,'HRV (Average)','Pulse Rate / Respiratory Rate / HRV',as.numeric)
  df$`Pulse (mean)` <- somfind(data,'Pulse Rate (Average)','Pulse Rate / Respiratory Rate / HRV',as.numeric)
  df$`Pulse (min)` <- somfind(data,'Pulse Rate (Lowest)','Pulse Rate / Respiratory Rate / HRV',as.numeric)
  df$`Pulse (max)` <- somfind(data,'Pulse Rate (Highest)','Pulse Rate / Respiratory Rate / HRV',as.numeric)
  
  return(data.frame(df,check.names = F))
}

# Sleep metrics Riss
sleepMetrics_CG <- function(data) {
  df <- list()
  # start date
  df$`Start date` <- sDate <- somfind(data,'Study Date','Referrer',function(x) as.Date(x,format = '%d/%m/%Y'))
  # start recording
  sr <- somfind(data,'Start Recording','Referrer',function(x) as.POSIXct(x, format = '%H:%M:%S', tz = 'GMT'))
  if (as.numeric(julian(sr)) - as.numeric(Sys.Date()) < 0.5) {
    df$`Start date` <- sDate <- sDate - 1
  }
  df$`Start recording` <- sr <- as.POSIXct(sub("\\S+", sDate, sr), tz = 'GMT')
  # end recording
  er <- somfind(data,'Stop Recording','Referrer',function(x) as.POSIXct(x, format = '%H:%M:%S'))
  df$`End recording` <- er <- as.POSIXct(sub("\\S+", sDate+1, er), tz = 'GMT')
  # lights out/on
  df$`Lights out` <- sr + somfind(data,'Lights Out Time','Sleep',as.numeric)
  df$`Lights on` <- sr + somfind(data,'Lights On Time','Sleep',as.numeric)
  # sleep onset/offset
  df$`Sleep onset` <- sr + somfind(data,'Sleep Onset','Sleep',as.numeric)
  df$`Sleep offset` <- sr + somfind(data,'Sleep Offset','Sleep',as.numeric)
  # total sleep time
  df$`Total sleep time (min)` <- somfind(data,'Total sleep time (TST)','Sleep',function(x) as.numeric(x) %>% `/`(60))
  # sleep latency
  df$`Sleep latency (min)` <- somfind(data,'Sleep Latency','Sleep',function(x) as.numeric(x) %>% `/`(60))
  df$`Sleep latency to 10 min (min)` <- somfind(data,' Latency (to 10 min sleep)','Sleep',function(x) as.numeric(x) %>% `/`(60))
  df$`REM latency (min)` <- somfind(data,' REM latency','Sleep',function(x) as.numeric(x) %>% `/`(60))
  # WASO
  df$`WASO (min)` <- somfind(data,'Wake after sleep onset (WASO)','Sleep',function(x) as.numeric(x) %>% `/`(60))
  # sleep efficiency
  df$`Sleep efficiency (%)` <- somfind(data,'Sleep efficiency','Sleep',as.numeric)
  # sleep availability
  df$`Sleep availability time (min)` <- somfind(data,'Total sleep period','Sleep',function(x) as.numeric(x) %>% `/`(60))
  df$`Time available for sleep (min)` <- somfind(data,'Time available for sleep','Sleep',function(x) as.numeric(x) %>% `/`(60))
  df$`Total recording time (min)` <- somfind(data,'Total Recording Time (TRT)','Referrer',function(x) as.numeric(x) %>% `/`(60))
  # sleep stages
  # time (mins)
  stages <- c('N1 Sleep Time','N2 Sleep Time','N1 Sleep Time','N3 Sleep Time','REM Sleep Time','NREM Sleep Time',
              'Unsure Time')
  cnames <- c('N1 (min)','N2 (min)','N1/N2 (min)','N3 (min)','REM (min)','NREM (min)','Unscored (min)')
  for (n in 1:length(stages)) {
    df[[cnames[n]]] <- somfind(data,stages[n],'Sleep',function(x) as.numeric(x) %>% `/`(60))
  }
  df$`N1/N2 (min)` <- df$`N1/N2 (min)` + df$`N2 (min)`
  # percentage
  stages <- c('Stage 1 / N1 %','Stage 2 / N2 %','Stage 1 / N1 %','Stage 3 / N3 %','REM sleep %','NREM sleep %')
  cnames <- c('N1 (%)','N2 (%)','N1/N2 (%)','N3 (%)','REM (%)','NREM (%)')
  for (n in 1:length(stages)) {
    df[[cnames[n]]] <- somfind(data,stages[n],'Sleep',as.numeric)
  }
  df$`N1/N2 (%)` <- df$`N1/N2 (%)` + df$`N2 (%)`
  # awakenings
  df$`Awakenings (total)` <- somfind(data,'Number of awakenings','Sleep',as.numeric)
  # pAHI
  df$`Resp events` <- somfind(data,'Qty Resp Events Total','Respiratory (PAT / Snore)',as.numeric)
  df$pAHI <- somfind(data,'pAHI Total','Respiratory (PAT / Snore)',as.numeric)
  # SpO2
  df$`Mean SpO2 (sleep)` <- somfind(data,'Average SpO2 (Sleep)','SpO2',as.numeric)
  df$`Min SpO2 (sleep)` <- somfind(data,'Lowest SpO2 (Sleep)','SpO2',as.numeric)
  df$`Mean desat (sleep)` <- somfind(data,'Average Desat (sleep)','SpO2',as.numeric)
  df$`Mean desat with resp events` <- somfind(data,'Average desaturation with respiratory events','SpO2',as.numeric)
  # HRV and pulse
  df$HRV <- somfind(data,'HRV (Average)','Pulse Rate / Respiratory Rate / HRV',as.numeric)
  df$`Pulse (mean)` <- somfind(data,'Pulse Rate (Average)','Pulse Rate / Respiratory Rate / HRV',as.numeric)
  df$`Pulse (min)` <- somfind(data,'Pulse Rate (Lowest)','Pulse Rate / Respiratory Rate / HRV',as.numeric)
  df$`Pulse (max)` <- somfind(data,'Pulse Rate (Highest)','Pulse Rate / Respiratory Rate / HRV',as.numeric)
  
  return(data.frame(df,check.names = F))
}

# Convert times to julian
t2j <- function(times, dates, se) {
  times <- sapply(c(1:length(times)), function(x)
    as.numeric(julian(as.POSIXct(times[x], origin = .Date(0), tz = 'GMT'), origin = dates[x])))
  if (grepl('s',se)) {
    if (any(times < 0.5)) {
      i <- which(times < 0.5)
      times[i] <- times[i] + 1
    }
  } else if (grepl('e',se)) {
    if (any(times > 0.5)) {
      i <- which(times > 0.5)
      times[i] <- times[i] - 1
    }
  }
  return(times)
}


# Subject names --------------------------------------------------------------------------
# extract subject names from RTF files
subNames <- function(filepath, filename) {
  subjects <- matrix(NA,length(filename),2)
  colnames(subjects) <- c('file','name')
  subjects[,1] <- basename(filename)
  for (i in 1:length(filename)) {
    # Import RTF
    temp <- rtf2df(filepath[i])
    # Participant ID
    subjects[i,2] <- temp$Study$`Metric Value`[which(temp$Study$`Metric Name` == 'URN')]
  }
  return(subjects)
}


# Load data: RTF -------------------------------------------------------------------------
# import, clean up and extract sleep metrics from RTF files
importFiles <- function(filepath, filename, mpcg) {
  fp <- dirname(filename[1])
  
  for (i in 1:length(filename)) {
    # Import RTF
    temp <- rtf2df(filepath[i])
    # Participant ID
    subject <- temp$Study$`Metric Value`[which(temp$Study$`Metric Name` == 'URN')]
    # Extract sleep metrics
    if (i > 1) {
      if (mpcg == 1) {
        sfMetrics <- rbind(sfMetrics,sleepMetrics_MP(temp))
      } else if (mpcg == 2) {
        sfMetrics <- rbind(sfMetrics,sleepMetrics_CG(temp))
      }
    } else {
      if (mpcg == 1) {
        sfMetrics <- sleepMetrics_MP(temp)
      } else if (mpcg == 2) {
        sfMetrics <- sleepMetrics_CG(temp)
      }
    }
  }
  data <- list('subject' = subject, 'sfMetrics' = sfMetrics, 'fp' = fp)
  return(data)
}


# Sleep report data ----------------------------------------------------------------------
sleepData <- function(data, rsn, nse) {
  # === Summary ===
  dex <- matrix(0,15,3)
  colnames(dex) <- c('Average','Minimum','Maximum')
  rownames(dex) <- c('Bed time','Wake-up time',
                     'Total time in bed','Sleep latency','Total sleep','WASO','Sleep efficiency (quality - %)',
                     'Awake (min)','N1/N2 (min)','N3 (min)','REM (min)','Awake (%)','N1/N2 (%)','N3 (%)','REM (%)')
  # replace start or end date with current date
  sr <- t2j(data$sfMetrics$`Start recording`,data$sfMetrics$`Start date`, 's')
  er <- t2j(data$sfMetrics$`End recording`, data$sfMetrics$`Start date` + 1, 'e')
  
  funs <- list(mean, min, max)
  dex[1,1:3] <- sapply(funs, function(fun, x) format(as.POSIXct(fun(x)*86400, origin = Sys.Date(), tz = 'GMT'), '%I:%M %p'), x = sr)
  dex[2,1:3] <- sapply(funs, function(fun, x) format(as.POSIXct(fun(x)*86400, origin = Sys.Date(), tz = 'GMT'), '%I:%M %p'), x = er)
  
  dex[3,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = data$sfMetrics$`Total recording time (min)`)
  dex[4,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = data$sfMetrics$`Sleep latency (min)`)
  dex[5,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = data$sfMetrics$`WASO (min)`)
  dex[6,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = data$sfMetrics$`Total sleep time (min)`)
  dex[7,1:3] <- sapply(funs, function(fun, x) round(fun(x, na.rm = T),1), x = data$sfMetrics$`Sleep efficiency (%)`)
  
  awake <- data$sfMetrics$`Time available for sleep (min)` - data$sfMetrics$`Total sleep time (min)`
  dex[8,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = data$sfMetrics$`N1/N2 (min)`)
  dex[9,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = data$sfMetrics$`N3 (min)`)
  dex[10,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = data$sfMetrics$`REM (min)`)
  dex[11,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = awake)
  awake <- awake / data$sfMetrics$`Total sleep time (min)` * 100
  dex[12,1:3] <- sapply(funs, function(fun, x) round(fun(x, na.rm = T)), x = data$sfMetrics$`N1/N2 (%)`)
  dex[13,1:3] <- sapply(funs, function(fun, x) round(fun(x, na.rm = T)), x = data$sfMetrics$`N3 (%)`)
  dex[14,1:3] <- sapply(funs, function(fun, x) round(fun(x, na.rm = T)), x = data$sfMetrics$`REM (%)`)
  dex[15,1:3] <- sapply(funs, function(fun, x) round(fun(x, na.rm = T)), x = awake)
  
  # === Data sheet ===
  sld <- data$sfMetrics
  sdates <- as.Date(sld$`Start date`, origin = .Date(0))
  edates <- as.Date(as.POSIXct(sld$`End recording`, origin = .Date(0)))
  # Convert bed and wake times to julian (fraction of day)
  sld$`Start recording` <- t2j(sld$`Start recording`, sdates, 's')
  sld$`End recording` <- t2j(sld$`End recording`, edates, 'e')
  sld$`Lights out` <- t2j(sld$`Lights out`, sdates, 's')
  sld$`Lights on` <- t2j(sld$`Lights on`, edates, 'e')
  sld$`Sleep onset` <- t2j(sld$`Sleep onset`, sdates, 's')
  sld$`Sleep offset` <- t2j(sld$`Sleep offset`, edates, 'e')
  
  sld$`Start date` <- as.Date(sld$`Start date`, origin = .Date(0))
  # account for bed time after midnight
  if (any(sld$`Lights on` < 0)) {
    i <- which(sld$`Lights on` < 0)
    sld$`Lights on`[i] <- sld$`Lights on`[i] + 1
  }
  
  # === Figure sheet ===
  slf <- list()
  slf$`Start date` <- sld$`Start date`
  # sleep latency, sleep duration, time awake,
  slf$`Sleep latency (h)` <- sld$`Sleep latency (min)`/60
  slf$`Sleep duration (h)` <- sld$`Total sleep time (min)`/60
  slf$`WASO (h)` <- sld$`WASO (min)`/60
  # Recommended sleep per night
  slf$`Recommended sleep per night (h)` <- rsn
  # Sleep efficiency and normal efficiency
  slf$`Sleep efficiency` <- sld$`Sleep efficiency (%)`
  slf$`Normal sleep efficiency (%)` <- nse
  # sleep pattern: bed time, wake-up time, average bed time, average wake-up time
  slf$`Bed time` <- sld$`Lights out`
  slf$`Wake-up time` <- (sld$`Lights on` + 1) - slf$`Bed time` # time difference between bed and wake time (for figure)
  slf$`Bed time avg` <- mean(slf$`Bed time`, na.rm = T)
  slf$`Wake-up time avg` <- mean(slf$`Wake-up time`, na.rm = T)
  
  slf$REM <- sld$`REM (min)`/(24*60)
  slf$`Deep sleep` <- sld$`N1/N2 (min)`/(24*60)
  slf$`Light sleep` <- sld$`N3 (min)`/(24*60)
  slf$`Time awake` <- (sld$`Time available for sleep (min)` - sld$`Total sleep time (min)`)/(24*60)
  
  slf <- data.frame(slf, check.names = F)
  slf[slf == 0] <- NA
  
  data$sld <- sld
  data$slf <- slf
  data$dex <- dex
  return(data)
}


# Export data to report template ---------------------------------------------------------
sleepReport <- function(data, srtemp) {
  # Load workbook
  wb <- loadWorkbook(srtemp)
  
  # Get worksheets
  sheets <- getSheets(wb)
  
  # Add name and date
  rows  <- getRows(sheets$`Report - Short`)
  cells <- getCells(rows)
  setCellValue(cells$`5.3`, gsub("_"," ",data$subject))
  setCellValue(cells$`6.3`, Sys.Date())
  
  # Add summary table
  addDataFrame(data$dex, sheets$Summary, startRow = 2, startColumn = 2, col.names = F, row.names = F)
  
  # Add data for figures
  addDataFrame(data$slf, sheets$Figures, startRow = 2, startColumn = 1, col.names = F, row.names = F, showNA = T)
  
  # Add data for analysis
  addDataFrame(data$sld, sheets$Data, startRow = 2, startColumn = 1, col.names = F, row.names = F, showNA = T)
  
  # Force cells to update
  wb$setForceFormulaRecalculation(T)
  
  return(wb)
}


# Create plots ---------------------------------------------------------------------------
# Plot sleep duration
durationPlot <- function(slf,rsn) {
  df <- cbind(slf[,c('Start date','Sleep latency (h)')],rep('Sleep latency (h)',nrow(slf)))
  colnames(df) <- c('Date','Hours','Phase')
  x <- cbind(slf[,c('Start date','Sleep duration (h)')],rep('Sleep duration (h)',nrow(slf)))
  colnames(x) <- c('Date','Hours','Phase')
  df <- rbind(df,x)
  x <- cbind(slf[,c('Start date','WASO (h)')],rep('WASO (h)',nrow(slf)))
  colnames(x) <- c('Date','Hours','Phase')
  df <- rbind(df,x)
  df$Hours[which(is.na(df$Hours))] <- 0
  df$Phase <- factor(df$Phase, levels = unique(df$Phase))
  
  ggplot(df) +
    geom_col(aes(x = Date, y = Hours, fill = Phase), position = position_stack(reverse = T), width = 1) + 
    scale_fill_manual(values = c('#8EB4E3','#002060','#376092')) +
    coord_cartesian(ylim = c(0,12),
                    xlim = c(min(df$Date)-0.75,max(df$Date)+0.75), expand = 0) +
    scale_y_continuous(breaks = c(0:12)) +
    ylab('Time (h)') +
    scale_x_date(date_breaks = "1 day", date_labels = "%d/%m") +
    geom_hline(aes(yintercept = rsn, linetype = 'Recommended sleep per night (h)'), colour = '#BE4B48', size = 1.25) +
    theme_minimal() +
    theme(legend.position = 'bottom',
          legend.title = element_blank(),
          axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1),
          axis.line = element_line(colour = 'black', size = 0.5, linetype = 'solid'),
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          axis.ticks = element_line(colour = 'black'))
}

# Plot sleep efficiency
efficiencyPlot <- function(slf,nse) {
  df <- slf[,c('Start date','Sleep efficiency')]
  df$`Sleep efficiency`[which(is.na(df$`Sleep efficiency`))] <- 0
  
  ggplot(df) +
    geom_col(aes(x = `Start date`, y = `Sleep efficiency`), fill = '#002060', width = 1) + 
    coord_cartesian(ylim = c(0,100),
                    xlim = c(min(df$`Start date`)-0.75,max(df$`Start date`)+0.75), expand = 0) +
    scale_y_continuous(breaks = seq(0,100,10)) +
    ylab('Efficiency (%)') +
    scale_x_date(date_breaks = "1 day", date_labels = "%d/%m") +
    geom_hline(aes(yintercept = nse, linetype = 'Normal sleep efficiency (%)'), colour = '#BE4B48', size = 1.25) +
    theme_minimal() +
    theme(legend.position = 'bottom', # 'none' to remove legend
          legend.title = element_blank(),
          axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1),
          axis.line = element_line(colour = 'black', size = 0.5, linetype = 'solid'),
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          axis.ticks = element_line(colour = 'black'))
}

# Plot sleep pattern
patternPlot <- function(slf,sld) {
  df <- cbind(slf[,c('Start date','Bed time')],rep('Bed time',nrow(slf)),rep(0,nrow(slf)))
  colnames(df) <- c('Date','Time','Night','Alpha')
  df$Time <- as.POSIXct(as.Date(df$Time, Sys.Date()-1))
  
  mwt <- as.POSIXct(as.Date(sld$`Lights on`, Sys.Date()))
  mwt <- as.POSIXct(as.numeric(mwt) - as.numeric(df$Time), origin = .Date(0))
  
  stages <- c('REM','Deep sleep','Light sleep','Time awake')
  for (s in stages) {
    x <- as.data.frame(as.numeric(slf[[s]] + slf$`Bed time`))
    x <- as.data.frame(cbind(slf$`Start date`,x,rep(s,nrow(slf)),rep(1,nrow(slf))))
    colnames(x) <- c('Date','Time','Night','Alpha')
    x$Time <- as.POSIXct(as.Date(x$Time, Sys.Date()))
    x$Time <- as.POSIXct(as.numeric(x$Time) - as.numeric(df$Time[df$Night == 'Bed time']) - 86400, origin = .Date(0))
    df <- rbind(df,x)
  }
  
  df[which(is.na(df))] <- 0
  df$Night <- factor(df$Night, levels = unique(df$Night))
  
  mbt <- mean(df$Time[which(df$Night == 'Bed time')])
  mwt <- as.POSIXct(as.numeric(mean(mwt)) + as.numeric(mbt), origin = .Date(0))
  
  ggplot(df) +
    geom_col(aes(x = Date, y = Time, fill = Night, alpha = Alpha), position = position_stack(reverse = T), width = 1) +
    scale_fill_manual(breaks = c('REM','Deep sleep','Light sleep','Time awake'),
                      values = c('green','#002060','#376092','#558ED5','#C6D9F1')) +
    scale_alpha_identity() +
    scale_y_datetime(labels = function(x) format(x, "%I:%M %p", tz = 'GMT'),
                     date_breaks = '2 hours', expand = c(0,0),
                     date_labels = "%I %p") +
    coord_cartesian(ylim = as.POSIXct(c(paste(Sys.Date()-1,'19:00:00'),paste(Sys.Date(),'11:00:00')), tz = 'GMT'),
                    xlim = c(min(df$Date)-0.75,max(df$Date)+0.75), expand = 0) +
    scale_x_date(date_breaks = "1 day", date_labels = "%d/%m") +
    geom_hline(aes(yintercept = mbt, linetype = 'Avg bed & wake-up time'), colour = '#BE4B48', size = 1.25) +
    geom_hline(yintercept = mwt, colour = '#BE4B48', linetype = 1, size = 1.25) +
    theme_minimal() +
    theme(legend.position = 'bottom',
          legend.title = element_blank(),
          axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1),
          axis.line.x = element_line(colour = 'black', size = 1, linetype = 'solid'),
          panel.grid.major.y = element_line(colour = '#D9D9D9', size = 0.75),
          panel.grid.minor.y = element_line(size = 0.75),
          panel.grid.major.x = element_blank(),
          panel.grid.minor.x = element_blank(),
          axis.ticks.x = element_line(colour = 'black'))
}


# # TEST -----------------------------------------------------------------------------------
# flist <- choose.files(caption = "Choose sleep tracking files")
# # srtemp <- choose.files(caption = "Select the sleep report template excel file")
# 
# subjects <- subNames(flist,flist)
# 
# for (s in unique(subjects[,2])) {
#   i <- grep(s,subjects[,2])
#   data <- importFiles(flist[i],flist[i],1)
#   data <- sleepData(data, 8, 85)
#   # sleepReport(data, srtemp)
# }
