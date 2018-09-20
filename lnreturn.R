options(scipen = 999)
library(openxlsx)

returncalculator <- function(filename){
  returnindex <- openxlsx::read.xlsx(filename,sheet ="Return Index")
  returnindexusd <- openxlsx::read.xlsx(filename,sheet ="Return Index - USD")
  deleted <- returnindex[,grepl("ERROR",colnames(returnindex))]
  deletedusd <- returnindexusd[,grepl("ERROR",colnames(returnindexusd))]
  returnindex <- returnindex[setdiff(colnames(returnindex),colnames(deleted))]
  returnindexusd <- returnindexusd[setdiff(colnames(returnindexusd),colnames(deletedusd))]
  nrows <- nrow(returnindex)
  ncols <- ncol(returnindex)
  returns <- data.frame(matrix(ncol = ncols-1,nrow = nrows))
  returnsusd <- data.frame(matrix(ncol = ncols-1,nrow = nrows))
  for (i in 2:nrows){  #calculates and writes log return data for each row
    returns[i,] <- log(returnindex[i,-1]) -log(returnindex[i-1,-1]) 
    returnsusd[i,] <- log(returnindexusd[i,-1]) -log(returnindexusd[i-1,-1]) 
  } # close for loop
  
  returns <- data.frame(format(as.Date(returnindex[,(1)], origin = "1899-12-30"),"%m") , format(as.Date(returnindex[,(1)], origin = "1899-12-30"),"%Y"), returns)
  colnames(returns)[1] <- "Month"
  colnames(returns)[2] <- "Year"
  colnames(returns)[3:(ncols+1)] <- colnames(returnindex)[(2:ncols)]
  returnsusd <- data.frame(format(as.Date(returnindexusd[,(1)], origin = "1899-12-30"),"%m") , format(as.Date(returnindexusd[,(1)], origin = "1899-12-30"),"%Y"), returnsusd[])
  colnames(returnsusd)[1] <- "Month"
  colnames(returnsusd)[2] <- "Year"
  colnames(returnsusd)[3:(ncols+1)] <- colnames(returnindexusd)[(2:ncols)]
  output <- gsub(".xlsx", "-Return.xlsx", filename)
  write.xlsx(returns, file = sprintf("./input.xlsx/%s", output)) 
  outputusd <- gsub(".xlsx", "-ReturnUSD.xlsx", filename)
  write.xlsx(returnsusd, file = sprintf("./input.xlsx/%s", outputusd)) 
} #close function

setwd ("~/Desktop/Data")
filelist <- "Estonia.xlsx"
for (filen in filelist){
  returncalculator(filen)
}
  
