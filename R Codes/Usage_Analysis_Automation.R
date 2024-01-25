# 1. SET AN END DATE

end_datå <- as.Date("2023-12-10")


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
    Ëèìèò = as.numeric(Ëèìèò),
    `Äúëæèìî äî çàíóëÿâàíå` = as.numeric(`Äúëæèìî äî çàíóëÿâàíå`),
    `Ðàçïîëàãàåìà ñóìà` = Ëèìèò - `Äúëæèìî äî çàíóëÿâàíå`
  )

  
desired_products <- c("AXI 2", 
                      "AXI 2-500",
                      "Visa Free ATM World",
                      "Visa Áÿëà Êàðòà - SC",
                      "Áÿëà Êàðòà", 
                      "Áÿëà Êàðòà – ñàìî ëèõâà", 
                      "Áÿëà Êàðòà 2",
                      "Áÿëà Êàðòà 3", 
                      "Áÿëà Êàðòà 3 - migrated",
                      "Áÿëà Êàðòà 4.2%",
                      "Áÿëà Êàðòà Gold - 10%",
                      "Áÿëà Êàðòà Gold - 5%",                                         
                      "ÈÀÌ")

filtered_data <- data[data$`Òåêóù èçäàòåë` == "EPS" &
                          data$Ïðîäóêò %in% desired_products &
                          data$`Ðàçïîëàãàåìà ñóìà` > 70 &
                          data$`Ñòàòóñ íà êàðòàòà` == "àêòèâíà" &
                          data$`Äíè çàáàâà` == 0, ]




columns_to_keep <- c("Òåêóù èçäàòåë",
                   "Ïðîäóêò",
                   "EasyClientNumber",
                   "Ëèìèò",
                   "Êëèåíò",
                   "ÅÃÍ",
                   "Òåëåôîí", 
                   "Ñòàòóñ íà êàðòàòà",
                   "Äúëæèìî äî çàíóëÿâàíå",
                   "Ðàçïîëàãàåìà ñóìà",
                   "Äíè çàáàâà",
                   "Äàòà íà ïîñëåäíî òåãëåíå íà ïàðè îò êàðòàòà")


data_csv <- filtered_data %>%
  select(all_of(columns_to_keep)) %>%
  rename(`Äàòà íà ïîñëåäíà òðàíçàêöèÿ` = `Äàòà íà ïîñëåäíî òåãëåíå íà ïàðè îò êàðòàòà`) %>%
  mutate(`Äàòà íà ïîñëåäíà òðàíçàêöèÿ` = format(as.POSIXct(`Äàòà íà ïîñëåäíà òðàíçàêöèÿ`, format = "%Y-%m-%d %H:%M"), "%d.%m.%Y"))



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
  rename(`Äàòà íà ïîäïèñâàíå íà äîãîâîð` = ContractDate) %>%
  mutate(`Äàòà íà ïîäïèñâàíå íà äîãîâîð` = format(as.Date(`Äàòà íà ïîäïèñâàíå íà äîãîâîð`), format = "%d.%m.%Y"))

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
  rename(`Äàòà íà ïîäïèñâàíå íà äîãîâîð` = ContractDate) %>%
  mutate(`Äàòà íà ïîäïèñâàíå íà äîãîâîð` = format(as.Date(`Äàòà íà ïîäïèñâàíå íà äîãîâîð`), format = "%d.%m.%Y"))

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
  rename(`Äàòà íà ïúðâà àêòèâàöèÿ íà êàðòàòà` = Date) %>%
  mutate(`Äàòà íà ïúðâà àêòèâàöèÿ íà êàðòàòà` = gsub("-", ".", `Äàòà íà ïúðâà àêòèâàöèÿ íà êàðòàòà`))

print(paste("Done in", round(difftime(Sys.time(), start_time, unit = "mins"), 2), "minutes"))

DBI::dbDisconnect(myc)



# Merge CSV file, first_activation_date, contract_date and consent_client

desired_order <- c("Òåêóù èçäàòåë",
                   "Ïðîäóêò",       
                   "EasyClientNumber",
                   "Ëèìèò",
                   "Êëèåíò",
                   "ÅÃÍ",
                   "Òåëåôîí", 
                   "Ñòàòóñ íà êàðòàòà",
                   "Äàòà íà ïúðâà àêòèâàöèÿ íà êàðòàòà",
                   "Äúëæèìî äî çàíóëÿâàíå",
                   "Ðàçïîëàãàåìà ñóìà",
                   "Äíè çàáàâà",
                   "Marketig_consent",
                   "SMS_Consent",
                   "Email_Consent",
                   "Email",
                   "Äàòà íà ïîäïèñâàíå íà äîãîâîð",
                   "Äàòà íà ïîñëåäíà òðàíçàêöèÿ"
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

# Email Anouncing

outlook <- COMCreate("Outlook.Application")
mail <- outlook$CreateItem(0)
mail[["Subject"]] <- "Closed Cards"


message <- "<html><body>"
message <- paste(message, "<p>Dear colleagues,</p>")
message <- paste(message, "<p>This email is automatically generated and contains information about added data.</p>")
message <- paste(message, "<p>Date and time of generation: ", format(Sys.time(), "%d %B %Y, %H:%M"), "</p>")
message <- paste(message, "<p>The Closed Cards data has been added successfully and can be found in the shared folder on PD.</p>")
message <- paste(message, "<p>J:/Product Development/Product Development_SHARED/BAR/Usage/Usage Analysis</p>")
message <- paste(message, "</body></html>")


mail[["HTMLBody"]] <- message

# mail[["To"]] <- "alexi.zein@gmail.com"

mail$Send()

cat("The notification email has been successfully sent.\n")



