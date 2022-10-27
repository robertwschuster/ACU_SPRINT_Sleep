#-----------------------------------------------------------------------------------------
#
# Load and analyze sleep tracking data (*.csv files)
# Robert Schuster (ACU SPRINT)
# August 2022
#
#-----------------------------------------------------------------------------------------


# clear environment
rm(list = ls())


# Libraries ------------------------------------------------------------------------------
if ("xlsx" %in% rownames(installed.packages()) == FALSE) {
  install.packages("xlsx")
}

if ("ggplot2" %in% rownames(installed.packages()) == FALSE) {
  install.packages("ggplot2")
}

library(xlsx)
library(ggplot2)


# Basic functions ------------------------------------------------------------------------
# convert minutes to hours:minutes
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
  funs <- list(mean, min, max)
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


# Load data ------------------------------------------------------------------------------
# select files to analyze
importFiles <- function(filepath,filename) {
  fp <- dirname(filename[1])
  subject <- unique(gsub('[[:digit:]]+', '', sub("([A-Za-z]+_[A-Za-z0-9]+).*", "\\1", basename(filename))))
  if (grepl("_",substr(subject,nchar(subject),nchar(subject)))) {
    subject <- gsub('[[:punct:]]', '', subject)
  }
  
  for (i in 1:length(filename)) {
    nCol <- max(count.fields(filepath[i], sep = ",")) # determine number of columns
    data <- read.csv(filepath[i], header = F, sep = ",", col.names = paste0("V",seq_len(nCol)), fill = T)
    data <- Filter(function(x)!all(is.na(x)), data) # remove columns that are all NAs
    data <- data[!apply(data == "", 1, all),] # remove empty rows
    
    # Statistics
    temp <- data[(which(grepl("Statistics",data[,1]))+1):(which(grepl("Marker/Score List",data[,1]))-1),]
    colnames(temp) <- temp[1,]
    temp <- temp[-c(1:2),]
    
    # Remove summaries from stats
    srows <- which(grepl("Summary",temp[,1]))
    
    # Merge files
    if (i > 1) {
      stats <- rbind(stats, temp[-srows,])
    } else {
      stats <- temp[-srows,]
    }
    
    stats <- stats[rowSums(stats[,3:ncol(stats)] == 'NaN') != ncol(stats)-2,] # remove rows that are all NaN
    stats <- stats[rowSums(is.na(stats[,3:ncol(stats)])) != ncol(stats)-2,] # remove rows that are all NA
    
    # convert columns to the right classes
    stats[c(2,7:18)] <- sapply(stats[c(2,7:18)],as.numeric)
    
    # Sort merged data
    if (i > 1) {
      # sort data based on 'Interval Type' and 'Start Date'
      stats <- stats[with(stats, order(as.factor(`Interval Type`), as.Date(`Start Date`,'%d/%m/%Y'))),]
      # remove duplicate rows based on 'Interval Type', 'Start Date' and 'End Date'
      stats <- stats[-which(duplicated(stats[,c('Interval Type','Start Date','End Date')])),]
      # number intervals
      for (f in unique(stats$`Interval Type`)) {
        r <- which(stats$`Interval Type` == f)
        stats$`Interval#`[r] <- c(1:length(r))
      }
    }
    
    # Statistics summary
    for (f in unique(stats$`Interval Type`)) {
      temp <- data.frame(matrix(NA,5,ncol(stats)))
      colnames(temp) <- colnames(stats)
      rownames(temp) <- paste(f,c('n','Minimum(n)','Maximum(n)','Average(n)','Std Dev(n-1)'))
      temp[1,c(7:18)] <- colSums(!is.na(stats[which(stats$`Interval Type` == f),c(7:18)]))
      temp[2,c(7:18)] <- apply(stats[which(stats$`Interval Type` == f),c(7:18)],2, 
                               function(x) ifelse(!all(is.na(x)), min(x, na.rm = T), NA))
      temp[3,c(7:18)] <- apply(stats[which(stats$`Interval Type` == f),c(7:18)],2,
                               function(x) ifelse(!all(is.na(x)), max(x, na.rm = T), NA))
      temp[4,c(7:18)] <- sapply(stats[which(stats$`Interval Type` == f),c(7:18)],
                                function(x) ifelse(!all(is.na(x)), mean(x, na.rm = T), NA))
      temp[5,c(7:18)] <- apply(stats[which(stats$`Interval Type` == f),c(7:18)],2,
                               function(x) ifelse(!all(is.na(x)), sd(x, na.rm = T), NA))
      
      if (f == unique(stats$`Interval Type`)[1]) {
        statsSum <- temp
      } else {
        statsSum <- rbind(statsSum, temp)
      }
    }
  }
  
  data <- list("stats" = stats, "statsSum" = statsSum, "subject" = subject, "fp" = fp)
  return(data)
}


# Sleep report data ----------------------------------------------------------------------
sleepData <- function(data, rsn, nse) {
  sleep <- which(data$stats$`Interval Type` == "SLEEP")
  ints <- unique(data$stats$`Interval#`[sleep])
  
  # Time in bed
  tib <- timeInBed(data$stats)
  
  # Sleep graphs data
  slf <- as.data.frame(matrix(0,length(ints),11))
  colnames(slf) <- c('Date','Bed time','Wake-up time','Sleep latency','Sleep duration','Time awake','Recommended sleep','Sleep efficiency',
                     'Normal efficiency','BT Avg','WT Avg')
  sld <- as.data.frame(matrix(0,length(ints),9))
  colnames(sld) <- c('Date','Bed time','Wake-up time','Time in bed','Sleep duration','Sleep latency','Sleep efficiency',
                     'WASO','#Wake bouts')
  
  dates <- as.Date(data$stats$`End Date`, '%d/%m/%Y')-1
  sdates <- vector()
  edates <- vector()
  st <- as.POSIXct(paste(data$stats$`Start Date`,data$stats$`Start Time`), '%d/%m/%Y %I:%M:%S %p', tz = 'GMT')
  et <- as.POSIXct(paste(data$stats$`End Date`,data$stats$`End Time`), '%d/%m/%Y %I:%M:%S %p', tz = 'GMT')
  
  for (i in ints) {
    r <- which(data$stats$`Interval Type` == 'SLEEP' & data$stats$`Interval#` == i)
    b <- which(data$stats$`Interval Type` == 'REST' & data$stats$`Interval#` == i & data$stats$`%Invalid SW` == 0)
    
    # Date
    slf[i,1] <- dates[r]
    sdates[i] <- as.Date(data$stats$`Start Date`[r], '%d/%m/%Y')
    edates[i] <- as.Date(data$stats$`End Date`[r], '%d/%m/%Y')
    # Bed time
    sld[i,2] <- st[b]
    # Wake-up time
    sld[i,3] <- et[b]
    # Sleep latency
    if ('Onset Latency' %in% colnames(data$stats)) {
      slf[i,4] <- sum(data$stats$`Onset Latency`[r], na.rm = T)/60 # hours
      sld[i,6] <- sum(data$stats$`Onset Latency`[r], na.rm = T) # minutes
    }
    # Sleep duration
    slf[i,5] <- sum(data$stats$`Sleep Time`[r], na.rm = T)/60 # hours
    sld[i,5] <- sum(data$stats$`Sleep Time`[r], na.rm = T) # minutes
    # Time in bed (minutes)
    sld[i,4] <- sum(data$stats$Duration[b], na.rm = T)
    # Time awake / light sleep (hours)
    slf[i,6] <- sum(data$stats$`Wake Time`[r], na.rm = T)/60
    # Sleep efficiency (%)
    slf[i,8] <- sld[i,7] <- sum(data$stats$Efficiency[r], na.rm = T)
    # WASO (minutes)
    if ('WASO' %in% colnames(data$stats)) {
      sld[i,8] <- sum(data$stats$`WASO`[r], na.rm = T)
    }
    # Number of wake bouts (#)
    if ('#Wake Bouts' %in% colnames(data$stats)) {
      sld[i,9] <- sum(data$stats$`#Wake Bouts`[r], na.rm = T)
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
  
  # Mean, min, max bed and wake times
  bt <- bwt(sld)
  
  data$sld <- sld
  data$slf <- slf
  data$dex <- rbind(bt,tib)
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
  setCellValue(cells$`8.2`, gsub("_"," ",data$subject))
  setCellValue(cells$`9.2`, Sys.Date())
  
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
  df <- cbind(slf[,c('Date','Sleep latency')],rep('Sleep latency',nrow(slf)))
  colnames(df) <- c('Date','Hours','Phase')
  x <- cbind(slf[,c('Date','Sleep duration')],rep('Sleep duration',nrow(slf)))
  colnames(x) <- c('Date','Hours','Phase')
  df <- rbind(df,x)
  x <- cbind(slf[,c('Date','Time awake')],rep('Time awake',nrow(slf)))
  colnames(x) <- c('Date','Hours','Phase')
  df <- rbind(df,x)
  df$Hours[which(is.na(df$Hours))] <- 0
  df$Phase <- factor(df$Phase, levels = unique(df$Phase))
  
  ggplot(df) +
    geom_col(aes(x = Date, y = Hours, fill = Phase), position = position_stack(reverse = T)) + 
    scale_fill_manual(values = c('#8EB4E3','#002060','#376092')) +
    coord_cartesian(ylim = c(0,12),
                    xlim = c(min(df$Date)-0.75,max(df$Date)+0.75), expand = 0) +
    scale_y_continuous(breaks = c(0:12)) +
    ylab('Time (h)') +
    scale_x_date(date_breaks = "2 days", date_labels = "%d/%m") +
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
  df <- slf[,c('Date','Sleep efficiency')]
  df$`Sleep efficiency`[which(is.na(df$`Sleep efficiency`))] <- 0
  
  ggplot(df) +
    geom_col(aes(x = Date, y = `Sleep efficiency`), fill = '#002060') + 
    coord_cartesian(ylim = c(0,100),
                    xlim = c(min(df$Date)-0.75,max(df$Date)+0.75), expand = 0) +
    scale_y_continuous(breaks = seq(0,100,10)) +
    ylab('Efficiency (%)') +
    scale_x_date(date_breaks = "2 days", date_labels = "%d/%m") +
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
  df <- cbind(slf[,c('Date','Bed time')],rep('Bed time',nrow(slf)),rep(0,nrow(slf)))
  colnames(df) <- c('Date','Time','Night','Alpha')
  df$Time <- as.POSIXct(as.Date(df$Time, Sys.Date()-1))
  
  x <- cbind(sld[,c('Date','Wake-up time')],rep('Wake-up time',nrow(slf)),rep(1,nrow(slf)))
  colnames(x) <- c('Date','Time','Night','Alpha')
  x$Time <- as.POSIXct(as.Date(x$Time, Sys.Date()))
  x$Time <- as.POSIXct(as.numeric(x$Time) - as.numeric(df$Time), origin = .Date(0))
  
  df <- rbind(df,x)
  df[which(is.na(df))] <- 0
  
  mbt <- mean(df$Time[which(df$Night == 'Bed time')])
  mwt <- as.POSIXct(as.numeric(mean(df$Time[which(df$Night == 'Wake-up time')])) + as.numeric(mbt), origin = .Date(0))
  
  ggplot(df) +
    geom_col(aes(x = Date, y = Time, fill = rev(Night), alpha = Alpha)) +
    scale_fill_manual(values = c('#002060','green')) +
    scale_alpha_identity() +
    scale_y_datetime(labels = function(x) format(x, "%I:%M %p", tz = 'GMT'),
                     date_breaks = '2 hours', expand = c(0,0),
                     date_labels = "%I %p") +
    coord_cartesian(ylim = as.POSIXct(c(paste(Sys.Date()-1,'19:00:00'),paste(Sys.Date(),'11:00:00')), tz = 'GMT'),
                    xlim = c(min(df$Date)-0.75,max(df$Date)+0.75), expand = 0) +
    scale_x_date(date_breaks = "2 days", date_labels = "%d/%m") +
    geom_hline(yintercept = mbt, colour = '#BE4B48', size = 1.25) +
    geom_hline(yintercept = mwt, colour = '#BE4B48', size = 1.25) +
    theme_minimal() +
    theme(legend.position = 'none',
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
# subjects <- unique(gsub('[[:digit:]]+', '', sub("([A-Za-z]+_[A-Za-z0-9]+).*", "\\1", basename(flist))))
# 
# for (s in 1:length(subjects)) {
#   if (grepl("_",substr(subjects[s],nchar(subjects[s]),nchar(subjects[s])))) {
#     subjects[s] <- gsub('[[:punct:]]', '', subjects[s])
#   }
# 
#   i <- grep(subjects[s],flist)
#   data <- importFiles(flist[i],flist[i])
#   data <- sleepData(data, 8, 85)
#   # sleepReport(data, srtemp)
# }
