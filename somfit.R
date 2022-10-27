#-----------------------------------------------------------------------------------------
#
# Load and analyze somfit sleep tracking data (*.docx files)
# Robert Schuster (ACU SPRINT)
# October 2022
#
#-----------------------------------------------------------------------------------------


# clear environment
rm(list = ls())


# Libraries ------------------------------------------------------------------------------
if ("striprtf" %in% rownames(installed.packages()) == F) {
  install.packages("striprtf")
}

if ("xlsx" %in% rownames(installed.packages()) == F) {
  install.packages("xlsx")
}

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
  df <- read_rtf(file, row_start = "*|", row_end = "", cell_end = "|")
  df <- gsub("\\*","", df) # remove asterix
  df <- lapply(df, function(x) strsplit(x, split = '\\|')[[1]])
  df <- lapply(df, function(x) x[2:length(x)]) # remove empty elements
  df <- df[lengths(df) == 4] # remove elements that didn't split properly (due to special characters, e.g., kiloohm)
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

# Sleep metrics
sleepMetrics <- function(data) {
  df <- as.data.frame(matrix(0,1,26))
  colnames(df) <- c('Start date','Start recording', 'End recording','Lights out','Lights on','Sleep onset','Sleep offset',
                    'Total sleep time (min)','WASO (min)','Sleep efficiency (min)',
                    'Sleep availability time (min)','Time available for sleep (min)','Total recording time (min)',
                    'N1 (min)','N2 (min)','N3 (min)','REM (min)','NREM (min)','Unscored (min)',
                    'N1 (%)','N2 (%)','N3 (%)','REM (%)','NREM (%)',
                    'Awakenings (total)','Awakenings (per hour)')
  # start date
  df[1,1] <- sDate <- as.Date(data$Referrer$`Metric Value`[which(data$Referrer$`Metric Name` == 'Study Date')], format = '%d/%m/%Y')
  # start recording
  sr <- as.POSIXct(data$Referrer$`Metric Value`[which(data$Referrer$`Metric Name` == 'Start Recording')], format = '%H:%M:%S')
  df[1,2] <- sr <- as.POSIXct(sub("\\S+", sDate, sr), tz = 'GMT')
  # end recording
  er <- as.POSIXct(data$Referrer$`Metric Value`[which(data$Referrer$`Metric Name` == 'Stop Recording')], format = '%H:%M:%S')
  df[1,3] <- er <- as.POSIXct(sub("\\S+", sDate+1, er), tz = 'GMT')
  # lights out/on
  df[1,4] <- sr + as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Lights Out Time')])
  df[1,5] <- sr + as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Lights On Time')])
  # sleep onset/offset
  df[1,6] <- sr + as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Sleep Onset')])
  df[1,7] <- sr + as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Sleep Offset')])
  # total sleep time
  df[1,8] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Total sleep time (TST)')])/60
  # WASO
  df[1,9] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Wake after sleep onset (WASO)')])/60
  # sleep efficiency
  df[1,10] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Sleep efficiency')])
  # sleep availability
  df[1,11] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Total sleep period')])/60
  df[1,12] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Time available for sleep')])/60
  df[1,13] <- as.numeric(data$Referrer$`Metric Value`[which(data$Referrer$`Metric Name` == 'Total Recording Time (TRT)')])/60
  # sleep stages
  # time (mins)
  stages <- c('N1 Sleep Time','N2 Sleep Time','N3 Sleep Time','REM Sleep Time','NREM Sleep Time','Unsure Time')
  for (n in 1:length(stages)) {
    df[1,13+n] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == stages[n])])/60
  }
  # percentage
  stages <- c('Stage 1 / N1 %','Stage 2 / N2 %','Stage 3 / N3 %','REM sleep %','NREM sleep %')
  for (n in 1:length(stages)) {
    df[1,19+n] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == stages[n])])
  }
  # awakenings
  df[1,25] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Number of awakenings')])
  df[1,26] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Awakenings Index')])
  
  return(df)
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


# Load data: RTF -------------------------------------------------------------------------
# select files to analyze
flist <- choose.files(caption = "Choose Somfit RTF files")
fp <- dirname(flist[1])

# sfData <- list()
sfMetrics <- list()
for (f in 1:length(flist)) {
  # Import RTF
  temp <- rtf2df(flist[f])
  # Participant ID
  s <- temp$Study$`Metric Value`[which(temp$Study$`Metric Name` == 'URN')]
  # sDate <- as.Date(temp$Referrer$`Metric Value`[which(temp$Referrer$`Metric Name` == 'Study Date')], format = '%d/%m/%Y')
  # sfData[[s]][[as.character(sDate)]] <- temp[c('Study','Referrer','Sleep')]
  if (s %in% names(sfMetrics)) {
    sfMetrics[[s]] <- rbind(sfMetrics[[s]],sleepMetrics(temp))
  } else {
    sfMetrics[[s]] <- sleepMetrics(temp)
  }
}
rm(f,s,temp)

subjects <- names(sfMetrics)

for (s in subjects) {
  # Summary
  dex <- matrix(0,10,3)
  colnames(dex) <- c('Average','Minimum','Maximum')
  rownames(dex) <- c('Bed time','Wake-up time',
                     'Total sleep','N1','N2','N3','REM','NREM',
                     'Sleep efficiency (quality - %)','WASO')
  sr <- as.POSIXct(sub("\\S+", Sys.Date(), as.POSIXct(sfMetrics[[s]]$`Start recording`, origin = .Date(0), tz = 'GMT')), tz = 'GMT')
  er <- as.POSIXct(sub("\\S+", Sys.Date(), as.POSIXct(sfMetrics[[s]]$`End recording`, origin = .Date(0), tz = 'GMT')), tz = 'GMT')
  
  funs <- list(mean, min, max)
  dex[1,1:3] <- sapply(funs, function(fun, x) format(fun(x), '%I:%M %p'), x = sr)
  dex[2,1:3] <- sapply(funs, function(fun, x) format(fun(x), '%I:%M %p'), x = er)
  
  dex[3,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = sfMetrics[[s]]$`Total sleep time`)
  dex[4,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = sfMetrics[[s]]$tN1)
  dex[5,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = sfMetrics[[s]]$tN2)
  dex[6,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = sfMetrics[[s]]$tN3)
  dex[7,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = sfMetrics[[s]]$tREM)
  dex[8,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = sfMetrics[[s]]$tNREM)
  dex[9,1:3] <- sapply(funs, function(fun, x) fun(x, na.rm = T), x = sfMetrics[[s]]$`Sleep efficiency`)
  dex[10,1:3] <- sapply(funs, function(fun, x) hm(fun(x, na.rm = T)), x = sfMetrics[[s]]$WASO)
  
  # Data sheet
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
  
  # Export to excel sheet
  write.xlsx(dex, file = paste0(fp,'/',s,'_somfit_data.xlsx'), sheetName = 'Summary')
  write.xlsx(sld, file = paste0(fp,'/',s,'_somfit_data.xlsx'), sheetName = 'Data', row.names = F, append = T)
}









# UNUSED ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# sfMetrics[[s]]$`Start recording` <- as.POSIXct(sfMetrics[[s]]$`Start recording`, origin = .Date(0))
# sfMetrics[[s]]$`End recording` <- as.POSIXct(sfMetrics[[s]]$`End recording`, origin = .Date(0))
# sfMetrics[[s]]$`Lights out` <- as.POSIXct(sfMetrics[[s]]$`Lights out`, origin = .Date(0))
# sfMetrics[[s]]$`Lights on` <- as.POSIXct(sfMetrics[[s]]$`Lights on`, origin = .Date(0))
# sfMetrics[[s]]$`Sleep onset` <- as.POSIXct(sfMetrics[[s]]$`Sleep onset`, origin = .Date(0))
# sfMetrics[[s]]$`Sleep offset` <- as.POSIXct(sfMetrics[[s]]$`Sleep offset`, origin = .Date(0))
# 
# # library(pdftools)
# 
# # https://stackoverflow.com/questions/44141160/recognize-pdf-table-using-r
# 
# 
# # Load data: PDF
# flist2 <- choose.files(caption = "Choose Somfit PDF files")
# x <- pdf_text(flist2[1][1])[2]
# 
# y <- strsplit(x, split = '\n')[[1]]
# z <- strsplit(x, split = '\r')[[1]]
# 
# tab <- extract_tables(flist2[1][1], pages = 2, area = c())
# 
# 
# # time
# df[1,13] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'N1 Sleep Time')])/60
# df[1,14] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'N2 Sleep Time')])/60
# df[1,15] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'N3 Sleep Time')])/60
# df[1,16] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'REM Sleep Time')])/60
# df[1,17] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'NREM Sleep Time')])/60
# df[1,18] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Unsure Time')])/60
# # percentage
# df[1,19] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Stage 1 / N1 %')])
# df[1,20] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Stage 2 / N2 %')])
# df[1,21] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Stage 3 / N3 %')])
# df[1,22] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'REM Sleep %')])
# df[1,23] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'NREM Sleep %')])


