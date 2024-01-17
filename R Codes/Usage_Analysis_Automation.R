# 1. SET AN END DATE

end_datе <- as.Date("2023-12-10")


# 2. ADD LINK FOR THE CARD REPORT

input_path <- "J:/Reports BG/Card_Report_BG/2023/12.2023/ASPxGridViewCards_11.12.2023.csv"

# 3. ADD OUTPUT LINK

output_path = "J:/BAR/R Scripts/R_scripts_PD_reports/Usage_Analisys_Automation/Output/Usage_10.12.2023.xlsx"

# 4. WAIT FOR THE PROGRAM.. 

# Importing Libraries 

library(readxl)
library(tidyverse)
library(odbc)
library(openxlsx)

# Template Loading 

template_path <- "J:/BAR/R Scripts/R_scripts_PD_reports/Usage_Analisys_Automation/Input/Usage.xlsx" 
workbook <- loadWorkbook(template_path)


# CSV Filtering

data <- read_delim(input_path, delim = ";", 
                   locale = locale(encoding = "cp1251"), 
                   show_col_types = FALSE)

data <- data %>% 
  slice(-2)


data <- data %>%
  mutate(
    Лимит = as.numeric(Лимит),
    `Дължимо до зануляване` = as.numeric(`Дължимо до зануляване`),
    `Разполагаема сума` = Лимит - `Дължимо до зануляване`
  )

  
desired_products <- c("AXI 2", 
                      "AXI 2-500",
                      "Visa Free ATM World",
                      "Visa Бяла Карта - SC",
                      "Бяла Карта", 
                      "Бяла Карта – само лихва", 
                      "Бяла Карта 2",
                      "Бяла Карта 3", 
                      "Бяла Карта 3 - migrated",
                      "Бяла Карта 4.2%",
                      "Бяла Карта Gold - 10%",
                      "Бяла Карта Gold - 5%",                                         
                      "ИАМ")

filtered_data <- data[data$`Текущ издател` == "EPS" &
                          data$Продукт %in% desired_products &
                          data$`Разполагаема сума` > 70 &
                          data$`Статус на картата` == "активна" &
                          data$`Дни забава` == 0, ]




columns_to_keep <- c("Текущ издател",
                   "Продукт",
                   "EasyClientNumber",
                   "Лимит",
                   "Клиент",
                   "ЕГН",
                   "Телефон", 
                   "Статус на картата",
                   "Дължимо до зануляване",
                   "Разполагаема сума",
                   "Дни забава",
                   "Дата на последно теглене на пари от картата")


data_csv <- filtered_data %>%
  select(all_of(columns_to_keep)) %>%
  rename(`Дата на последна транзакция` = `Дата на последно теглене на пари от картата`) %>%
  mutate(`Дата на последна транзакция` = format(as.POSIXct(`Дата на последна транзакция`, format = "%Y-%m-%d %H:%M"), "%d.%m.%Y"))



# SQL Queries Loading


# Contract Date Loading
myc <- DBI::dbConnect(odbc::odbc(),
                      driver = "SQL Server",
                      server = "scorpio.smartitbg.int",
                      database = "BIsmartWCBG")

sql_query_1 <- read_file("J:/BAR/R Scripts/R_scripts_PD_reports/Usage_Analisys_Automation/SQL Querries/Contract_Date.sql")

start_time <- Sys.time()
contract_date <- DBI::dbGetQuery(myc, sql_query_1) %>%
  select(EasyClientNumber, ContractDate) %>%
  rename(`Дата на подписване на договор` = ContractDate) %>%
  mutate(`Дата на подписване на договор` = format(as.Date(`Дата на подписване на договор`), format = "%d.%m.%Y"))

print(paste("Done in", round(difftime(Sys.time(), start_time, unit = "mins"), 2), "minutes"))

DBI::dbDisconnect(myc)


# Consent Client Loading
myc <- DBI::dbConnect(odbc::odbc(),
                      driver = "SQL Server",
                      server = "scorpio.smartitbg.int",
                      database = "BIsmartWCBG")

sql_query_1 <- read_file("J:/BAR/R Scripts/R_scripts_PD_reports/Usage_Analisys_Automation/SQL Querries/Contract_Date.sql")

start_time <- Sys.time()
contract_date <- DBI::dbGetQuery(myc, sql_query_1) %>%
  select(EasyClientNumber, ContractDate) %>%
  rename(`Дата на подписване на договор` = ContractDate) %>%
  mutate(`Дата на подписване на договор` = format(as.Date(`Дата на подписване на договор`), format = "%d.%m.%Y"))

print(paste("Done in", round(difftime(Sys.time(), start_time, unit = "mins"), 2), "minutes"))

DBI::dbDisconnect(myc)



# First Activation Date Loading
myc <- DBI::dbConnect(odbc::odbc(),
                      driver = "SQL Server",
                      server = "scorpio.smartitbg.int",
                      database = "BIsmartWCBG")

sql_query_3 <- read_file("J:/BAR/R Scripts/R_scripts_PD_reports/Usage_Analisys_Automation/SQL Querries/FirstActivationDate.sql")

start_time <- Sys.time()
first_activation_date <- DBI::dbGetQuery(myc, sql_query_3) %>%
  select(EasyClientNumber, Date) %>%
  rename(`Дата на първа активация на картата` = Date) %>%
  mutate(`Дата на първа активация на картата` = gsub("-", ".", `Дата на първа активация на картата`))

print(paste("Done in", round(difftime(Sys.time(), start_time, unit = "mins"), 2), "minutes"))

DBI::dbDisconnect(myc)



# Merge CSV file, first_activation_date, contract_date and consent_client

desired_order <- c("Текущ издател",
                   "Продукт",       
                   "EasyClientNumber",
                   "Лимит",
                   "Клиент",
                   "ЕГН",
                   "Телефон", 
                   "Статус на картата",
                   "Дата на първа активация на картата",
                   "Дължимо до зануляване",
                   "Разполагаема сума",
                   "Дни забава",
                   "Marketig_consent",
                   "SMS_Consent",
                   "Email_Consent",
                   "Email",
                   "Дата на подписване на договор",
                   "Дата на последна транзакция"
)


data_csv <- data_csv %>%
  merge(first_activation_date, by.x = "EasyClientNumber", by.y = "EasyClientNumber", all.x = TRUE) %>%
  merge(contract_date, by.x = "EasyClientNumber", by.y = "EasyClientNumber", all.x = TRUE) %>%
  merge(consent_client, by.x = "EasyClientNumber", by.y = "EasyClientNumber", all.x = TRUE) %>%
  select(all_of(desired_order)) %>%
  mutate(across(everything(), ~ifelse(is.na(.), "#N/A", .)))


# Data Export

writeData(workbook, sheet = "Sheet1", x = data_csv, startRow = 2, startCol = 1, colNames = FALSE)
saveWorkbook(workbook, file = output_path, overwrite = TRUE)





