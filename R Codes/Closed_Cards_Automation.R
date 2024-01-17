# 1. SET A TIME PERIOD

start_date <- as.Date('2023-12-01')
end_date <- as.Date('2023-12-30')

# 2. ADD OUTPUT LINK

output <- "J:/BAR/R Scripts/R_scripts_PD_reports/Closed_Cards_Automation/Output/ClosedCards.xlsx"

# WAIT FOR THE PROGRAM..

# Importing Libraries 


library(tidyverse)
library(readxl)
library(odbc)
library(openxlsx)

#library(RDCOMClient)

# Template Loading 

input <- "J:/BAR/R Scripts/R_scripts_PD_reports/Closed_Cards_Automation/Input/ClosedCards.xlsx"
workbook <- loadWorkbook(input)


# SQL Queries Loading

server = "Scorpio"
database = "BIsmartWCBG"

conn <- dbConnect(odbc::odbc(), 
                  driver = "ODBC Driver 17 for SQL Server", 
                  server = server, 
                  database = database, 
                  Trusted_Connection = "Yes")

sql_query <- paste("
  DECLARE @StartDate DATE = '", start_date, "';
  DECLARE @EndDate DATE = '", end_date, "';

  SELECT 
    DimOff.ContractNumber AS EasyClientNumber,
    CONVERT(DATE, DimOff.DateClosed) AS Date,
    DimPr.Name AS Product,
    DimCR.Code AS CloseReason,
    COUNT(*) AS Count 
  FROM dwh.DimOffers AS DimOff
  JOIN dwh.DimOffCloseReason AS DimCR ON DimCR.CloseReasonSK = DimOff.CloseReasonSK
  JOIN dwh.DimProduct AS DimPr ON DimPr.ProductSK = DimOff.ProductSK
  WHERE DimCR.CloseReasonSK BETWEEN 1 AND 3 
    AND CONVERT(DATE, DimOff.DateClosed) BETWEEN @StartDate AND @EndDate
  GROUP BY DimOff.ContractNumber, DimPr.Name, DimCR.Code, CONVERT(DATE, DimOff.DateClosed);
", collapse = "")

conn <- dbConnect(odbc::odbc(), 
                  driver = "ODBC Driver 17 for SQL Server", 
                  server = server, 
                  database = database, 
                  Trusted_Connection = "Yes")

data_sql <- dbGetQuery(conn, sql_query) %>%
  mutate(CloseReason = ifelse(CloseReason %in% c("Внесено над лимита", "Недостиг до зануляване"), "VoluntaryChurn", "Cession"),
         Date = format(as.Date(Date), "%d.%m.%Y"))


# Data Export
workbook <- loadWorkbook(input)
writeData(workbook, sheet = "Sheet1", x = data_sql, startRow = 2, startCol = 1, colNames = FALSE)
saveWorkbook(workbook, file = output, overwrite = TRUE)


