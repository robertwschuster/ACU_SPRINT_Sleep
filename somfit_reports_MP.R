#-----------------------------------------------------------------------------------------
#
# Load and analyze somfit sleep tracking data (*.rtf files)
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

##########################################################################################


# Libraries ------------------------------------------------------------------------------
if ("striprtf" %in% rownames(installed.packages()) == F) {
  install.packages("striprtf")
}

if ("xlsx" %in% rownames(installed.packages()) == F) {
  install.packages("xlsx")
}

library(tidyverse)
library(striprtf)
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
  # if (rv == '-' || rv == '- ' || identical(rv,integer(0))) {
  if (rv == '-' || rv == '- ' || length(rv) == 0) {
    rv <- NA
  } else {
    if (!is.null(proc)) {
      rv <- rv %>% proc
    }
  }
  return(rv)
}

# Sleep metrics
sleepMetrics <- function(data) {
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
  sr <- somfind(data,'Start Recording','Referrer',function(x) as.POSIXct(x, format = '%H:%M:%S'))
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


# Load sleep report excel template -------------------------------------------------------
srtemp <- "C:/Users/s4548745/Desktop/ACU/ACU_SPRINT_Sleep/Report templates/Report template_somfit_MP.xlsx"

if (!file.exists(srtemp)) {
  srtemp <- choose.files(caption = "Select the sleep report template excel file")
}


# Load data: RTF -------------------------------------------------------------------------
# select files to analyze
flist <- choose.files(caption = "Choose Somfit RTF files")
fp <- dirname(flist[1])


# Extract sleep metrics ------------------------------------------------------------------
sfMetrics <- list()
for (f in 1:length(flist)) {
  # Import RTF
  temp <- rtf2df(flist[f])
  # Participant ID
  s <- temp$Study$`Metric Value`[which(temp$Study$`Metric Name` == 'URN')]
  # Extract sleep metrics
  if (s %in% names(sfMetrics)) {
    sfMetrics[[s]] <- rbind(sfMetrics[[s]],sleepMetrics(temp))
  } else {
    sfMetrics[[s]] <- sleepMetrics(temp)
  }
}
rm(f,s,temp)

subjects <- names(sfMetrics)


# Export data to report template ---------------------------------------------------------
for (s in subjects) {
  # === Summary ===
  dex <- matrix(0,15,3)
  colnames(dex) <- c('Average','Minimum','Maximum')
  rownames(dex) <- c('Bed time','Wake-up time',
                     'Total time in bed','Sleep latency','Total sleep','WASO','Sleep efficiency (quality - %)',
                     'Awake (min)','N1/N2 (min)','N3 (min)','REM (min)','Awake (%)','N1/N2 (%)','N3 (%)','REM (%)')
  # replace start or end date with current date
  sr <- as.POSIXct(sub("\\S+", Sys.Date(), as.POSIXct(sfMetrics[[s]]$`Start recording`, origin = .Date(0), tz = 'GMT')), tz = 'GMT')
  er <- as.POSIXct(sub("\\S+", Sys.Date(), as.POSIXct(sfMetrics[[s]]$`End recording`, origin = .Date(0), tz = 'GMT')), tz = 'GMT')
  
  funs <- list(mean, min, max)
  dex[1,1:3] <- sapply(funs, function(fun, x) format(fun(x), '%I:%M %p'), x = sr)
  dex[2,1:3] <- sapply(funs, function(fun, x) format(fun(x), '%I:%M %p'), x = er)
  
  dex[3,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = sfMetrics[[s]]$`Total recording time (min)`)
  dex[4,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = sfMetrics[[s]]$`Sleep latency (min)`)
  dex[5,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = sfMetrics[[s]]$`WASO (min)`)
  dex[6,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = sfMetrics[[s]]$`Total sleep time (min)`)
  dex[7,1:3] <- sapply(funs, function(fun, x) round(fun(x, na.rm = T),1), x = sfMetrics[[s]]$`Sleep efficiency (%)`)
  
  awake <- sfMetrics[[s]]$`Time available for sleep (min)` - sfMetrics[[s]]$`Total sleep time (min)`
  dex[8,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = sfMetrics[[s]]$`N1/N2 (min)`)
  dex[9,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = sfMetrics[[s]]$`N3 (min)`)
  dex[10,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = sfMetrics[[s]]$`REM (min)`)
  dex[11,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = awake)
  awake <- awake / sfMetrics[[s]]$`Total sleep time (min)` * 100
  dex[12,1:3] <- sapply(funs, function(fun, x) round(fun(x, na.rm = T)), x = sfMetrics[[s]]$`N1/N2 (%)`)
  dex[13,1:3] <- sapply(funs, function(fun, x) round(fun(x, na.rm = T)), x = sfMetrics[[s]]$`N3 (%)`)
  dex[14,1:3] <- sapply(funs, function(fun, x) round(fun(x, na.rm = T)), x = sfMetrics[[s]]$`REM (%)`)
  dex[15,1:3] <- sapply(funs, function(fun, x) round(fun(x, na.rm = T)), x = awake)
  
  # === Data sheet ===
  sld <- sfMetrics[[s]]
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
  
  # === Export to report template ===
  # Load workbook
  wb <- loadWorkbook(srtemp)
  # Get worksheets
  sheets <- getSheets(wb)
  # Add name and date
  rows  <- getRows(sheets$`Report - Short`)
  cells <- getCells(rows)
  setCellValue(cells$`5.3`, gsub("_"," ",s))
  setCellValue(cells$`6.3`, Sys.Date())
  # Add summary table
  addDataFrame(dex, sheets$Summary, startRow = 2, startColumn = 2, col.names = F, row.names = F)
  # Add data for figures
  addDataFrame(slf, sheets$Figures, startRow = 2, startColumn = 1, col.names = F, row.names = F, showNA = T)
  # Add data for analysis
  addDataFrame(sld, sheets$Data, startRow = 2, startColumn = 1, col.names = F, row.names = F, showNA = T)
  # Force cells to update
  wb$setForceFormulaRecalculation(T)
  # Save the workbook
  saveWorkbook(wb, paste0(fp,'/',s,'-somfit_report_MP.xlsx'))
}


# UNUSED ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# reverse julian date
# as.POSIXct((sld$`Sleep onset`[1]*86400), origin = sdates[1] , tz = 'GMT')
