install.packages("gmodels")


library(tidyverse)
library(gmodels)




input <- "C:/Users/aleksandar.dimitrov/Desktop/INFOLITICA/IFRS 9 ÏÎÄÃÎÒÎÂÊÀ/employee-data.csv"

data <- read.csv(input)

head(data)



CrossTable(data$title)







