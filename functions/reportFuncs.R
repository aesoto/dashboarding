makeReports <- function(attributionData) {
  fundData <- getFundData()
  # for testing - define these variables and ender the loop as this specific case, lines 9-11,then 14-
  account <- '235000'
  day <- names(attributionData)[1]
  
  output <- list()
  for(day in names(attributionData)) {
    attributionPacket <- attributionData[[day]]
    printDate <- as.Date(day, format='%Y.%m.%d')
    printDate <- format(printDate, format='%m/%d/%Y')
    
    for(account in names(attributionPacket)){
      reportDate <- day
      attributionPanel <- attributionPacket[[account]]
      fundName <- getFundName(account, fundData)
      accountTitle <- paste0(account, ' - ', fundName, ' - Top Down Attribution Summary')
      
      # procure metrics for dashboard
      summaryData <- getSummaryData(attributionPanel)
      totalReturn <- getTotalReturn(summaryData, printDate)
      attribution <- getAttribution(summaryData, printDate)
      allocation <- getAllocation(attributionPanel)
      categoryReturns <- getCategoryReturns(attributionPanel, printDate)
      positionsData <- processPositions(attributionPanel, account, reportDate)
      
      # record day's data for trailing calculations, retain historical record for further calculation
      attributionData <- storeAttributionData(attributionPanel, account, reportDate)
      attributionHistory <- storeAttributionMetric(account, attribution, reportDate)
      returnHistory <- storeTotalReturn(account, totalReturn, reportDate)
      categoryHistory <- storeCategoryReturns(account, categoryReturns, reportDate)
      
      # procure trailing metrics
      trailingMetrics <- getTrailingMetrics(reportDate, attributionHistory, returnHistory, categoryHistory)
      
      # print metrics onto template
      publishData(allocation, totalReturn, attribution,
                  categoryReturns, positionsData, account, reportDate, trailingMetrics)
    }
  }
}

getFundData <- function() {
  file.name <- 'S:/Unit10230/CTI_CMS Project/Scripts for Data Push/Development/Fund List v3.csv'
  fundData <- read.csv(file.name, stringsAsFactors = FALSE)[ ,c('Account', 'Description')]
  fundData <- fundData[complete.cases(fundData),]
  return(fundData)
}

getFundName <- function(account, fundData) {
  fundName <- fundData$Description[fundData$Account==account]
  return(fundName)
}

getSummaryData <- function(attributionPanel) {
  summaryData <- attributionPanel[attributionPanel$Index=='Totals',]
  summaryData <- summaryData[c('NAV', 'PortfolioCTR', 'otherAdj', 'priceAdj', 'AttribCTR', 'BenchCTR', 'basis',
                               'Level1Attr', 'Level2Attr', 'Level3Attr', 'Level4Attr', 'Level5Attr',
                               'SelectionAttr', 'actvRet'
  )]
  summaryData$Date <- reportDate
  summaryData <- select(summaryData, Date, everything())
  return(summaryData)
}

getTotalReturn <- function(summaryData, printDate) {
  returnData <- select(summaryData, PortfolioCTR:actvRet)
  returnData <- as.data.frame(t(returnData))
  rownames(returnData) <- NULL
  returnData <- data.frame(returnData, stringsAsFactors = FALSE)
  returnData <- head(returnData, 6)
  names(returnData)[1]<-eval(printDate)
  returnData <- round(returnData, 6)
  return(returnData)
}

getAttribution <- function(summaryData, printDate) {
  returnData <- select(summaryData, PortfolioCTR:actvRet)
  returnData <- as.data.frame(t(returnData))
  rownames(returnData) <- NULL
  returnData <- data.frame(returnData, stringsAsFactors = FALSE)
  attribData <- tail(returnData, 7)
  names(attribData)[1]<-eval(printDate)
  attribData <- round(attribData, 6)
  return(attribData)
}

getAllocation <- function(attributionPanel) {
  allocation <- attributionPanel[c('Level2', 'Level3', 'PortWt', 'BenchWt')]
  allocation <- allocation[!is.na(allocation$Level2) & is.na(allocation$Level3),]
  allocation$Level3 <- NULL
  allocation$`Over/Under` <- allocation$PortWt - allocation$BenchWt
  US <- allocation[allocation$Level2=='US',]
  US$Level2 <- 'US Equity'
  allocation <- allocation[!allocation$Level2=='US',]
  allocation$Level2[allocation$Level2=='Intl'] <- 'International Equity'
  allocation <- rbind(US, allocation)  
  totals <- data.frame(Level2='Total',
                       PortWt=sum(allocation$PortWt),
                       BenchWt=sum(allocation$BenchWt)
  )
  totals$`Over/Under` <- 0
  allocation <- rbind(allocation, totals)
  names(allocation) <- c('Allocation', 'Portfolio', 'Benchmark', 'Over/Under')
  allocation$Allocation <- NULL
  allocation <- round(allocation, 6)
  return(allocation)
}

getCategoryReturns <- function(attributionPanel, printDate) {
  assetClasses <- c('US', 'Intl', 'Fixed Income', 'Other')
  categoryReturns <- data.frame(stringsAsFactors = FALSE)
  for(assetClass in assetClasses) {
    if(assetClass=='Other') {
      assetData <- attributionPanel[attributionPanel$Level1==assetClass & is.na(attributionPanel$Level2),]
      assetData <- assetData[complete.cases(assetData[, 10:28]),]
      assetData <- select(assetData, PortfolioCTR1:basis)
      assetData <- as.data.frame(t(assetData))
      rownames(assetData) <- NULL
      assetData <- data.frame(assetData, stringsAsFactors = FALSE)
      assetData <- round(assetData, 6)
      names(assetData)[1] <- eval(printDate)
    } else {
      assetData <- attributionPanel[attributionPanel$Level2==assetClass & is.na(attributionPanel$Level3),]
      assetData <- assetData[complete.cases(assetData[, 10:28]),]
      assetData <- select(assetData, PortfolioCTR1:basis)
      assetData <- as.data.frame(t(assetData))
      rownames(assetData) <- NULL
      assetData <- data.frame(assetData, stringsAsFactors = FALSE)
      assetData <- round(assetData, 6)
      names(assetData)[1] <- eval(printDate)
    }
    categoryReturns <- rbind(categoryReturns, assetData)
  }
  return(categoryReturns)
}

processPositions <- function(attributionPanel, account, reportDate) {
  allData <- storeHoldingsData(attributionPanel, account, reportDate)
  MTDReturn <- calculateMTDReturn(allData, reportDate)
  topHoldings <- getTopHoldings(attributionPanel)
  bottomHoldings <- getBottomHoldings(attributionPanel)
  positionsData <- list(bestDay=topHoldings, worstDay=bottomHoldings,
                        bestMTD=MTDReturn$bestMTD, worstMTD=MTDReturn$worstMTD)
  return(positionsData)
}

storeHoldingsData <- function(attributionPanel, account, reportDate){
  calendarDate <- as.Date(reportDate, '%Y.%m.%d')
  fileDate <- format(calendarDate, '%Y.%m')
  account.file.name <- paste0(fileDate, '_Historical_Holdings_', account, '.csv')
  output.path <- paste0('S:/Unit10230/CTI_CMS Project/Top Down Attribution/FoFs/Data/Holdings/', account.file.name)
  categories <- c('US', 'Intl', 'Fixed Income')
  
  holdingsData <- attributionPanel[attributionPanel$Selection!='',]
  holdingsData <- holdingsData[c('Level2', 'Selection', 'Sec_Des', 'PortRet', 'BenchRet', 'actvRet')]
  holdingsData$Date <- reportDate
  holdingsData <- select(holdingsData, Date, everything())
  
  holdingsData <- holdingsData[holdingsData$Level2 %in% categories,]
  holdingsData <- holdingsData[order(holdingsData$Selection),]
  holdingsData$PortRet <- 1 - holdingsData$PortRet
  holdingsData$BenchRet <- 1 - holdingsData$BenchRet
  holdingsData$actvRet <- 1 - holdingsData$actvRet
  
  portfolio <- holdingsData[c('Date', 'Level2', 'Selection', 'Sec_Des', 'PortRet')]
  colnames(portfolio)[colnames(portfolio)=='PortRet'] <- 'return'
  colnames(portfolio)[colnames(portfolio)=='Date'] <- 'date'
  portfolio$returnType <- 'portfolio'
  portfolio <- select(portfolio, returnType, everything())
  benchmark <- holdingsData[c('Date', 'Level2', 'Selection', 'Sec_Des', 'BenchRet')]
  colnames(benchmark)[colnames(benchmark)=='BenchRet'] <- 'return'
  colnames(benchmark)[colnames(benchmark)=='Date'] <- 'date'
  benchmark$returnType <- 'benchmark'
  benchmark <- select(benchmark, returnType, everything())
  active <- holdingsData[c('Date', 'Level2', 'Selection', 'Sec_Des', 'actvRet')]
  colnames(active)[colnames(active)=='actvRet'] <- 'return'
  colnames(active)[colnames(active)=='Date'] <- 'date'
  active$returnType <- 'active'
  active <- select(active, returnType, everything())
  
  totalData <- rbind(portfolio, benchmark, active)
  
  firstDay <- bizdays::getdate('first bizday', ref(calendarDate, ym='month'), 'mycal')
  if(calendarDate == firstDay) { 
    write.csv(file=output.path, x=totalData, row.names=FALSE)
    #prevDate <- bizdays::offset(calendarDate, -1, 'mycal')
    #prevMonth <- format(prevDate, '%Y.%m')
    #account.file.name <- paste0(prevMonth, '_Historical_Holdings_', account, '.csv')
    #output.path <- paste0('S:/Unit10230/CTI_CMS Project/Top Down Attribution/FoFs/Data/Holdings/', account.file.name)
    #fileData <- fread(output.path)
    allData <- totalData
  } else {
    fileData <- read.csv(file=output.path)
    #fileData <- fread(output.path)
    fileData <- data.frame(fileData, stringsAsFactors = FALSE)
    fileData <- fileData[fileData$date!=reportDate,]
    allData <- rbind(totalData, fileData)
    write.csv(file=output.path, x=allData, row.names=FALSE)
  }
  return(allData)
}

calculateMTDReturn <- function(allData, reportDate) {
  date <- as.Date(reportDate, format='%Y.%m.%d')
  prevDate <- bizdays::getdate('first bizday', ref(date, ym='month'), 'mycal')
  returnDates <- bizdays::bizseq(prevDate, date, 'mycal')
  returnDates <- format(returnDates, '%Y.%m.%d')
  
  positions <- allData[allData$returnType=='active',]
  positions <- positions[positions$date %in% returnDates,]
  
  returns <- data.frame(Selection=as.character(),
                        Sec_Des=as.character(),
                        return=as.numeric())
  for(day in returnDates) {
    temp <- positions[positions$date==day,]
    temp <- temp[c('Selection', 'Sec_Des', 'return')]
    if(nrow(returns)==0) { 
      returns <- temp 
    } else {
      returns <- returns %>% 
        full_join(temp, by=c('Selection', 'Sec_Des'))
      returns <- returns %>%
        mutate_if(is.numeric,~ ifelse(is.na(.), 1, .))
    }
  }
  temp <- select(returns, Selection, Sec_Des)
  temp1 <- select(returns, -Selection, -Sec_Des)
  temp1 <- temp1 %>%
    mutate('actvRet' = Reduce('*', .))
  temp1 <- temp1['actvRet']
  returns <- cbind(temp, temp1)
  
  winnerPanel <- getTopMTDHoldings(returns, returnDates, allData)
  loserPanel <- getBottomMTDHoldings(returns, returnDates, allData)
  mtdData <- list(bestMTD=winnerPanel, worstMTD=loserPanel)
  return(mtdData)
}

getTopMTDHoldings <- function(returns, returnDates, allData){
  winners <- returns[order(-returns$actvRet),]
  winners <- head(winners, 10)
  winnerData <- allData[allData$Selection %in% winners$Selection,]
  
  positions <- winnerData[winnerData$returnType=='portfolio',]
  positions <- positions[positions$date %in% returnDates,]
  portfolioReturns <- data.frame(Selection=as.character(),
                                 Sec_Des=as.character(),
                                 return=as.numeric())
  for(day in returnDates) {
    temp <- positions[positions$date==day,]
    temp <- temp[c('Selection', 'Sec_Des', 'return')]
    if(nrow(portfolioReturns)==0) { 
      portfolioReturns <- temp 
    } else {
      portfolioReturns <- portfolioReturns %>% 
        full_join(temp, by=c('Selection', 'Sec_Des'))
      portfolioReturns <- portfolioReturns %>%
        mutate_if(is.numeric,~ ifelse(is.na(.), 1, .))
    }
  }
  temp <- select(portfolioReturns, Selection, Sec_Des)
  temp1 <- select(portfolioReturns, -Selection, -Sec_Des)
  temp1 <- temp1 %>%
    mutate('PortRet' = Reduce('*', .))
  temp1 <- temp1['PortRet']
  portfolioReturns <- cbind(temp, temp1)
  
  winnerPanel <- winners %>%
    full_join(portfolioReturns, by=c('Selection', 'Sec_Des'))
  
  
  positions <- winnerData[winnerData$returnType=='benchmark',]
  positions <- positions[positions$date %in% returnDates,]
  benchmarkReturns <- data.frame(Selection=as.character(),
                                 Sec_Des=as.character(),
                                 return=as.numeric())
  for(day in returnDates) {
    temp <- positions[positions$date==day,]
    temp <- temp[c('Selection', 'Sec_Des', 'return')]
    if(nrow(benchmarkReturns)==0) { 
      benchmarkReturns <- temp 
    } else {
      benchmarkReturns <- benchmarkReturns %>% 
        full_join(temp, by=c('Selection', 'Sec_Des'))
      benchmarkReturns <- benchmarkReturns %>%
        mutate_if(is.numeric,~ ifelse(is.na(.), 1, .))
    }
  }
  temp <- select(benchmarkReturns, Selection, Sec_Des)
  temp1 <- select(benchmarkReturns, -Selection, -Sec_Des)
  temp1 <- temp1 %>%
    mutate('BenchRet' = Reduce('*', .))
  temp1 <- temp1['BenchRet']
  benchmarkReturns <- cbind(temp, temp1)
  
  winnerPanel <- winnerPanel %>%
    full_join(benchmarkReturns, by=c('Selection', 'Sec_Des'))
  
  winnerPanel <- winnerPanel[c('Selection', 'Sec_Des', 'PortRet', 'BenchRet', 'actvRet')]
  winnerPanel$Selection <- NULL
  winnerPanel$PortRet <- winnerPanel$PortRet - 1
  winnerPanel$BenchRet <- winnerPanel$BenchRet - 1
  winnerPanel$actvRet <- winnerPanel$actvRet - 1
  names(winnerPanel) <- c('Top Contributors - MTD', 'MTD', 'Bench Return', 'Relative Contr.')
  winnerPanel <- winnerPanel %>% mutate_if(is.numeric, list(~round(.,6)))
  return(winnerPanel)
}

getBottomMTDHoldings <- function(returns, returnDates, allData){
  losers <- returns[order(returns$actvRet),]
  losers <- head(losers, 10)
  loserData <- allData[allData$Selection %in% losers$Selection,]
  
  positions <- loserData[loserData$returnType=='portfolio',]
  positions <- positions[positions$date %in% returnDates,]
  portfolioReturns <- data.frame(Selection=as.character(),
                                 Sec_Des=as.character(),
                                 return=as.numeric())
  for(day in returnDates) {
    temp <- positions[positions$date==day,]
    temp <- temp[c('Selection', 'Sec_Des', 'return')]
    if(nrow(portfolioReturns)==0) { 
      portfolioReturns <- temp 
    } else {
      portfolioReturns <- portfolioReturns %>% 
        full_join(temp, by=c('Selection', 'Sec_Des'))
      portfolioReturns <- portfolioReturns %>%
        mutate_if(is.numeric,~ ifelse(is.na(.), 1, .))
    }
  }
  temp <- select(portfolioReturns, Selection, Sec_Des)
  temp1 <- select(portfolioReturns, -Selection, -Sec_Des)
  temp1 <- temp1 %>%
    mutate('PortRet' = Reduce('*', .))
  temp1 <- temp1['PortRet']
  portfolioReturns <- cbind(temp, temp1)
  
  loserPanel <- losers %>%
    full_join(portfolioReturns, by=c('Selection', 'Sec_Des'))
  
  
  positions <- loserData[loserData$returnType=='benchmark',]
  positions <- positions[positions$date %in% returnDates,]
  benchmarkReturns <- data.frame(Selection=as.character(),
                                 Sec_Des=as.character(),
                                 return=as.numeric())
  for(day in returnDates) {
    temp <- positions[positions$date==day,]
    temp <- temp[c('Selection', 'Sec_Des', 'return')]
    if(nrow(benchmarkReturns)==0) { 
      benchmarkReturns <- temp 
    } else {
      benchmarkReturns <- benchmarkReturns %>% 
        full_join(temp, by=c('Selection', 'Sec_Des'))
      benchmarkReturns <- benchmarkReturns %>%
        mutate_if(is.numeric,~ ifelse(is.na(.), 1, .))
    }
  }
  temp <- select(benchmarkReturns, Selection, Sec_Des)
  temp1 <- select(benchmarkReturns, -Selection, -Sec_Des)
  temp1 <- temp1 %>%
    mutate('BenchRet' = Reduce('*', .))
  temp1 <- temp1['BenchRet']
  benchmarkReturns <- cbind(temp, temp1)
  
  loserPanel <- loserPanel %>%
    full_join(benchmarkReturns, by=c('Selection', 'Sec_Des'))
  
  loserPanel <- loserPanel[c('Selection', 'Sec_Des', 'PortRet', 'BenchRet', 'actvRet')]
  loserPanel$Selection <- NULL
  loserPanel$PortRet <- loserPanel$PortRet - 1
  loserPanel$BenchRet <- loserPanel$BenchRet - 1
  loserPanel$actvRet <- loserPanel$actvRet - 1
  names(loserPanel) <- c('Top Detractors - MTD', 'MTD', 'Bench Return', 'Relative Contr.')
  loserPanel <- loserPanel %>% mutate_if(is.numeric, list(~round(.,6)))
  return(loserPanel)
}

getTopHoldings <- function(attributionPanel) {
  categories <- c('US', 'Intl', 'Fixed Income', 'Other')
  holdingsData <- attributionPanel[attributionPanel$Selection!='',]
  holdingsData <- holdingsData[c('Level2', 'Selection', 'Sec_Des', 'PortRet', 'BenchRet', 'actvRet')]
  holdingsData <- holdingsData[holdingsData$Level2 %in% categories,]
  holdingsData <- holdingsData[order(holdingsData$Selection),]
  holdingsData$PortRet <- 1 - holdingsData$PortRet
  holdingsData$BenchRet <- 1 - holdingsData$BenchRet
  holdingsData$actvRet <- 1 - holdingsData$actvRet
  
  winners <- holdingsData[order(-holdingsData$actvRet),]
  winners <- head(winners, 10)
  winners <- winners[c('Sec_Des', 'PortRet', 'BenchRet', 'actvRet')]
  winners$PortRet <- winners$PortRet - 1
  winners$BenchRet <- winners$BenchRet - 1
  winners$actvRet <- winners$actvRet - 1
  names(winners) <- c('Top Contributors - Daily', 'Daily Return', 'Bench Return', 'Relative Contr.')
  winners <- winners %>% 
    mutate_if(is.numeric, list(~round(.,6)))
  return(winners)
}

getBottomHoldings <- function(attributionPanel) {
  categories <- c('US', 'Intl', 'Fixed Income', 'Other')
  holdingsData <- attributionPanel[attributionPanel$Selection!='',]
  holdingsData <- holdingsData[c('Level2', 'Selection', 'Sec_Des', 'PortRet', 'BenchRet', 'actvRet')]
  holdingsData <- holdingsData[holdingsData$Level2 %in% categories,]
  holdingsData <- holdingsData[order(holdingsData$Selection),]
  holdingsData$PortRet <- 1 - holdingsData$PortRet
  holdingsData$BenchRet <- 1 - holdingsData$BenchRet
  holdingsData$actvRet <- 1 - holdingsData$actvRet
  
  losers <- holdingsData[order(holdingsData$actvRet),]
  losers <- head(losers, 10)
  losers <- losers[c('Sec_Des', 'PortRet', 'BenchRet', 'actvRet')]
  losers$PortRet <- losers$PortRet - 1
  losers$BenchRet <- losers$BenchRet - 1
  losers$actvRet <- losers$actvRet - 1
  names(losers) <- c('Bottom Contributors - Daily', 'Daily Return', 'Bench Return', 'Relative Contr.')
  losers <- losers %>% mutate_if(is.numeric, list(~round(.,6)))
  return(losers)
}

storeAttributionData <- function(attributionPanel, account, reportDate) {
  # Stores historical attribution data for multi-period calculation
  accountData <- attributionPanel
  account.file.name <- paste0(account, '_Historical_Attribution_TopLine.csv')
  output.path <- paste0('S:/Unit10230/CTI_CMS Project/Top Down Attribution/FoFs/Data/', account.file.name)
  
  summaryData <- accountData[accountData$Index=='Totals',]
  summaryData <- summaryData[c('NAV', 'PortfolioCTR', 'otherAdj', 'priceAdj', 'AttribCTR', 'BenchCTR', 'basis',
                               'Level1Attr', 'Level2Attr', 'Level3Attr', 'Level4Attr', 'Level5Attr',
                               'SelectionAttr', 'actvRet'
  )]
  NAV <- summaryData$NAV[1]
  summaryData <- summaryData %>% mutate_if(is.numeric, list(~1-.))
  summaryData$NAV[1] <- NAV
  summaryData$Date <- reportDate
  summaryData <- select(summaryData, Date, everything())
  
  fileData <- read.csv(file=output.path)
  fileData <- data.frame(fileData, stringsAsFactors = FALSE)
  fileData <- fileData[fileData$Date!=reportDate,]
  fileData <- rbind(fileData, summaryData)
  fileData <- fileData[order(fileData$Date),]
  write.csv(file=output.path, x=fileData, row.names=FALSE)
  return(fileData)
}

storeAttributionMetric <- function(account, attribution, reportDate) {
  # Stores historical attribution data for multi-period calculation
  account.file.name <- paste0(account, '_Historical_Attribution.csv')
  output.path <- paste0('S:/Unit10230/CTI_CMS Project/Top Down Attribution/FoFs/Data/', account.file.name)
  
  returnData <- data.frame(Date=reportDate,
                           Level1=1-attribution[1,],
                           Level2=1-attribution[2,],
                           Level3=1-attribution[3,],
                           Level4=1-attribution[4,],
                           Level5=1-attribution[5,],
                           Selection=1-attribution[6,],
                           TotalActvRet=1-attribution[7,])
  fileData <- read.csv(file=output.path)
  fileData <- data.frame(fileData, stringsAsFactors = FALSE)
  fileData <- fileData[fileData$Date!=reportDate,]
  fileData <- rbind(fileData, returnData)
  fileData <- fileData[order(fileData$Date),]
  write.csv(file=output.path, x=fileData, row.names=FALSE)
  return(fileData)
}

storeTotalReturn <- function(account, totalReturn, reportDate) {
  # Stores historical return data for multi-period calculation
  account.file.name <- paste0(account, '_Historical_Return.csv')
  output.path <- paste0('S:/Unit10230/CTI_CMS Project/Top Down Attribution/FoFs/Data/', account.file.name)
  
  returnData <- data.frame(Date=reportDate,
                           Ret=1-totalReturn[1,],
                           otherAdj=1-totalReturn[2,],
                           priceAdj=1-totalReturn[3,],
                           attribRet=1-totalReturn[4,],
                           bench=1-totalReturn[5,],
                           basis=1-totalReturn[6,])
  fileData <- read.csv(file=output.path)
  fileData <- data.frame(fileData, stringsAsFactors = FALSE)
  fileData <- fileData[fileData$Date!=reportDate,]
  fileData <- rbind(fileData, returnData)
  fileData <- fileData[order(fileData$Date),]
  write.csv(file=output.path, x=fileData, row.names=FALSE)
  return(fileData)
}

storeCategoryReturns <- function(account, categoryReturns, reportDate) {
  # Stores historical return data for multi-period calculation
  account.file.name <- paste0(account, '_Historical_Category_Returns.csv')
  output.path <- paste0('S:/Unit10230/CTI_CMS Project/Top Down Attribution/FoFs/Data/', account.file.name)
  
  categoryData <- data.frame(Date=reportDate,
                             USEquity=1-categoryReturns[1,],
                             USPriceAdj=1-categoryReturns[2,],
                             USAttribRet=1-categoryReturns[3,],
                             USBench=1-categoryReturns[4,],
                             USBasis=1-categoryReturns[5,],
                             intlEquity=1-categoryReturns[6,],
                             intlPriceAdj=1-categoryReturns[7,],
                             intlAttribRet=1-categoryReturns[8,],
                             intlBench=1-categoryReturns[9,],
                             intlBasis=1-categoryReturns[10,],
                             FixedIncome=1-categoryReturns[11,],
                             FIPriceAdj=1-categoryReturns[12,],
                             FIAttribRet=1-categoryReturns[13,],
                             FIBench=1-categoryReturns[14,],
                             FIBasis=1-categoryReturns[15,],
                             Other=1-categoryReturns[16,],
                             otherPriceAdj=1-categoryReturns[17,],
                             otherAttribRet=1-categoryReturns[18,],
                             otherBench=1-categoryReturns[19,],
                             otherBasis=1-categoryReturns[20,]
  )
  fileData <- read.csv(file=output.path)
  fileData <- data.frame(fileData, stringsAsFactors = FALSE)
  fileData <- fileData[fileData$Date!=reportDate,]
  fileData <- rbind(fileData, categoryData)
  fileData <- fileData[order(fileData$Date),]
  write.csv(file=output.path, x=fileData, row.names=FALSE)
  return(fileData)
}

publishData <- function(allocation, totalReturn, attribution,
                        categoryReturns, positionsData, account, reportDate, trailingMetrics) {
  account.file.name <- paste0(reportDate, '_Top_Down_Attribution_Report_', account, '.xlsx')
  template.name <- paste0('Top Down Attribution Template ',account, '.xlsx')
  template.path <- paste0('S:/Unit10230/CTI_CMS Project/Top Down Attribution/FoFs/Dashboards/', template.name)
  output.path <- paste0('S:/Unit10230/CTI_CMS Project/Top Down Attribution/FoFs/Dashboards/', account.file.name)
  
  bestDay <- positionsData$bestDay
  worstDay <- positionsData$worstDay
  bestMTD <- positionsData$bestMTD
  worstMTD <- positionsData$worstMTD
  
  wb <- loadWorkbook(template.path, create = FALSE)
  setStyleAction(wb,XLC$"STYLE_ACTION.NONE")
  
  publishDataGroups <- c('Summary', 'US', 'Intl', 'Fixed Income', 'Other')
  
 for(ws_publishData in publishDataGroups) {
  
  publishDataGroup <- readWorksheet(wb, ws_publishData)
  writeWorksheet(wb, allocation, sheet=ws_publishData, startRow=3, startCol=3)
  writeWorksheet(wb, totalReturn, sheet=ws_publishData, startRow=12, startCol=3)
  writeWorksheet(wb, attribution, sheet=ws_publishData, startRow=20, startCol=3)
  writeWorksheet(wb, categoryReturns, sheet=pws_publishData, startRow=30, startCol=3)
  
  writeWorksheet(wb, bestDay, sheet=ws_publishData, startRow=3, startCol=14)
  writeWorksheet(wb, worstDay, sheet=ws_publishData, startRow=15, startCol=14)
  
  writeWorksheet(wb, bestMTD, sheet=ws_publishData, startRow=29, startCol=14)
  writeWorksheet(wb, worstMTD, sheet=ws_publishData, startRow=41, startCol=14)
  
  # publish trailing metrics
  categories <- c('return', 'attribution', 'category')
  
  for(period in names(trailingMetrics)){
    dataPacket <- trailingMetrics[[period]]
    for(category in names(dataPacket)){
      reportData <- dataPacket[[category]]
      if(category=='return') {
        row <- 12
      } else if (category=='attribution') {
        row <- 20
      } else {
        row <- 30
      }
      if(period=='5Day') {
        column <- 4
      } else if(period=='MTD') {
        column <- 5
      } else if(period=='QTD') {
        column <- 6
      } else if(period=='3M') {
        column <- 7 
      } else if(period=='YTD') {
        column <- 8
      } else {
        column <- 9
      }
      writeWorksheet(wb, reportData, sheet=ws_publishData, startRow=row, startCol=column)
    }
  }
  saveWorkbook(wb, output.path)
 }
}
  
  
#period choices are: 5day, MTD, QTD, 3M, YTD, 1Y
getDays <- function(reportDate, period) {
  calendarDate <- as.Date(reportDate, '%Y.%m.%d')
  if(period=='5Day') {
    startDate <- offset(calendarDate, -6)
    periodDays <- bizseq(startDate, calendarDate, 'mycal')
  } else if(period=='MTD') {
    startDate <- bizdays::getdate('first bizday', ref(calendarDate, ym='month'), 'mycal')
    periodDays <- bizseq(startDate, calendarDate, 'mycal')
  } else if(period=='QTD') {
    month <- format(calendarDate, '%m')
    year <- format(calendarDate, '%Y')
    firstQuarter <- c('01', '02', '03')
    secondQuarter <- c('04', '05', '06')
    thirdQuarter <- c('07', '08', '09')
    fourthQuarter <- c('10', '11', '12')
    
    if(month %in% firstQuarter) {
      startDate <- bizdays::getdate('first bizday', ref(calendarDate, ym='year'), 'mycal')
    } else if (month %in% secondQuarter) {
      refDate <- paste0(year,'.04.15')
      refDate <- as.Date(refDate, '%Y.%m.%d')
      startDate <- bizdays::getdate('first bizday', ref(refDate, ym='month'), 'mycal')
    } else if (month %in% thirdQuarter) {
      refDate <- paste0(year,'.07.15')
      refDate <- as.Date(refDate, '%Y.%m.%d')
      startDate <- bizdays::getdate('first bizday', ref(refDate, ym='month'), 'mycal')
    } else {
      refDate <- paste0(year,'.09.15')
      refDate <- as.Date(refDate, '%Y.%m.%d')
      startDate <- bizdays::getdate('first bizday', ref(refDate, ym='month'), 'mycal')
    }
    periodDays <- bizseq(startDate, calendarDate, 'mycal')
  } else if(period=='3M') {
    startDate <- offset(calendarDate, -91)
    periodDays <- bizseq(startDate, calendarDate, 'mycal')
  } else if(period=='YTD') {
    startDate <- bizdays::getdate('first bizday', ref(calendarDate, ym='year'), 'mycal')
    periodDays <- bizseq(startDate, calendarDate, 'mycal')
  } else if(period=='1Y') {
    startDate <- offset(calendarDate, -366)
    periodDays <- bizseq(startDate, calendarDate, 'mycal')
  }
  return(periodDays)
}

calculateReturns <- function(reportDate, period, historyData) {
  #print(historyData[1,])
  periodDays <- getDays(reportDate, period)
  periodDays <- format(periodDays, format='%Y.%m.%d')
  returnData <- historyData[historyData$Date %in% periodDays,]
  
  if('NAV' %in% names(returnData)) { returnData$NAV <- NULL }
  
  returnCategories <- names(returnData)
  returnCategories <- returnCategories[ returnCategories != 'Date']
  
  temp <- NULL
  for(category in returnCategories) {
    temp[category] <- prod(returnData[category])
  }
  
  temp <- as.data.frame(t(temp))
  temp <- as.data.frame(t(temp))
  rownames(temp) <- NULL
  temp <- data.frame(temp, stringsAsFactors = FALSE)
  names(temp)[1]<-eval(period)
  temp[1] <- temp[1]-1
  return(temp)
}


getTrailingMetrics <- function(reportDate, attributionHistory, returnHistory, categoryHistory) {
  periods <- c('5Day', 'MTD', 'QTD', '3M', 'YTD', '1Y')
  dataFields <- c('attribution', 'return', 'category')
  trailingData <- list()
  
  for(period in periods) {
    periodData <- list()
    for(dataField in dataFields) {
      
      result <- calculateReturnMetric(reportDate, dataField, period, attributionHistory, returnHistory, categoryHistory)
      #print(result)
      periodData[[dataField]] <- result 
    }
    trailingData[[period]] <- periodData
  }
  return(trailingData)
}

calculateReturnMetric <- function(reportDate, dataField, period, attributionHistory, returnHistory,
                                  categoryHistory) {
  if(dataField=='attribution') {
    result <- calculateReturns(reportDate, period, attributionHistory)
    #print('Using attribution data')
  } else if (dataField=='return') {
    result <- calculateReturns(reportDate, period, returnHistory)
    #print('Using return data')
  } else {
    result <- calculateReturns(reportDate, period, categoryHistory)
    #print('Using category data')
  }
  return(result)
}
