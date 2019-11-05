getData <- function(dates) {
  # ToDo: once source data is stored, convert this to cycle through both types of accounts
  accountType <- 'FoFs/Archive/'
  #accountType <- 'Funds/Archive/'
  
  data.path <- 'S:/Unit10230/CTI_CMS Project/Top Down Attribution/'
  data.path <- paste0(data.path, accountType)
  
  panels <- list()
  for(day in dates$days) {
    if(accountType=='FoFs/Archive/') { file.name <- paste0(day,' FoF Top Down Attribution.xlsx') }
    if(accountType=='Funds/Archive/') { file.name <- paste0(day,' Top Down Attribution.xlsx') }
    file <- paste0(data.path, file.name)
    sheets <- excel_sheets(file)
    sheets <- sheets[sheets != 'metrics']
    accounts <- list()
    for(sheet in sheets){
      accountPanel <- read_excel(file, sheet=sheet)
      accountPanel <- select(accountPanel, -contains('Val'))
      accountPanel <- data.frame(accountPanel, stringsAsFactors = FALSE)
      accounts[[sheet]] <- accountPanel
    }
    panels[[day]] <- accounts
  }
  return(panels)
}
