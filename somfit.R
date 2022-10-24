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
if ("striprtf" %in% rownames(installed.packages()) == FALSE) {
  install.packages("striprtf")
}

library(striprtf)

library(pdftools)

# https://stackoverflow.com/questions/44141160/recognize-pdf-table-using-r



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
  df <- as.data.frame(matrix(0,1,25))
  colnames(df) <- c('Total sleep time','WASO','Sleep efficiency',
                    'Start recording', 'End recording','Lights out','Lights on','Sleep onset','Sleep offset',
                    'Sleep availability time','Time available for sleep','Total recording time',
                    'tN1','tN2','tN3','tREM','tNREM','tUnscored',
                    'pN1','pN2','pN3','pREM','pNREM',
                    'nAwakenings','iAwakenings')
  
  df[1,1] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Total sleep time (TST)')])/60
  df[1,2] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Wake after sleep onset (WASO)')])/60
  df[1,3] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Sleep efficiency')])
  
  # start date
  sDate <- as.Date(data$Referrer$`Metric Value`[which(data$Referrer$`Metric Name` == 'Study Date')], format = '%d/%m/%Y')
  # start recording
  sr <- as.POSIXct(data$Referrer$`Metric Value`[which(data$Referrer$`Metric Name` == 'Start Recording')], format = '%H:%M:%S')
  df[1,4] <- sr <- replace(sr, T, sub("\\S+", sDate, sr))
  # end recording
  er <- as.POSIXct(data$Referrer$`Metric Value`[which(data$Referrer$`Metric Name` == 'Stop Recording')], format = '%H:%M:%S')
  df[1,5] <- er <- replace(er, T, sub("\\S+", sDate+1, er))
  # lights out/on
  df[1,6] <- sr + as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Lights Out Time')])
  df[1,7] <- sr + as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Lights On Time')])
  # sleep onset/offset
  df[1,8] <- sr + as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Sleep Onset')])
  df[1,9] <- sr + as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Sleep Offset')])
  # sleep availability
  df[1,10] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Total sleep period')])/60
  df[1,11] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Time available for sleep')])/60
  df[1,12] <- as.numeric(data$Referrer$`Metric Value`[which(data$Referrer$`Metric Name` == 'Total Recording Time (TRT)')])/60
  # sleep stages
  # time (mins)
  stages <- c('N1 Sleep Time','N2 Sleep Time','N3 Sleep Time','REM Sleep Time','NREM Sleep Time','Unsure Time')
  for (n in 1:length(stages)) {
    df[1,12+n] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == stages[n])])/60
  }
  # percentage
  stages <- c('Stage 1 / N1 %','Stage 2 / N2 %','Stage 3 / N3 %','REM sleep %','NREM sleep %')
  for (n in 1:length(stages)) {
    df[1,18+n] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == stages[n])])
  }
  # awakenings
  df[1,24] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Number of awakenings')])
  df[1,25] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Awakenings Index')])
  
  return(df)
}


# Load data: RTF -------------------------------------------------------------------------
# select files to analyze
flist <- choose.files(caption = "Choose Somfit RTF files")
fp <- dirname(flist[1])

sfData <- list()
sfMetrics <- list()
for (f in 1:length(flist)) {
  temp <- rtf2df(flist[f])

  # Participant ID
  s <- temp$Study$`Metric Value`[which(temp$Study$`Metric Name` == 'URN')]
  sDate <- as.Date(temp$Referrer$`Metric Value`[which(temp$Referrer$`Metric Name` == 'Study Date')], format = '%d/%m/%Y')
  
  sfData[[s]][[as.character(sDate)]] <- temp[c('Study','Referrer','Sleep')]
  if (s %in% names(sfMetrics)) {
    sfMetrics[[s]] <- rbind(sfMetrics[[s]],sleepMetrics(temp))
  } else {
    sfMetrics[[s]] <- sleepMetrics(temp)
  }
  
  sfMetrics[[s]]$`Start recording` <- as.POSIXct(sfMetrics[[s]]$`Start recording`, origin = .Date(0))
}

subjects <- names(sfData)






# total sleep period, time available for sleep and total recording time: use hm() after summary metrics (min, max, mean) 
# computation!





# UNUSED ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# Load data: PDF
flist2 <- choose.files(caption = "Choose Somfit PDF files")
x <- pdf_text(flist2[1][1])[2]

y <- strsplit(x, split = '\n')[[1]]
z <- strsplit(x, split = '\r')[[1]]

tab <- extract_tables(flist2[1][1], pages = 2, area = c())


# time
df[1,13] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'N1 Sleep Time')])/60
df[1,14] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'N2 Sleep Time')])/60
df[1,15] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'N3 Sleep Time')])/60
df[1,16] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'REM Sleep Time')])/60
df[1,17] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'NREM Sleep Time')])/60
df[1,18] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Unsure Time')])/60
# percentage
df[1,19] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Stage 1 / N1 %')])
df[1,20] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Stage 2 / N2 %')])
df[1,21] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'Stage 3 / N3 %')])
df[1,22] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'REM Sleep %')])
df[1,23] <- as.numeric(data$Sleep$`Metric Value`[which(data$Sleep$`Metric Name` == 'NREM Sleep %')])


