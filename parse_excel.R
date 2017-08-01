# Parse ND .xls files

library(readxl)
library(dplyr)
library(stringr)
library(readr)
library(tidyr)

files <- list.files(path = "../openelections-sources-nd", 
                    pattern = "20161108__nd__general__*", 
                    full.names = TRUE)
offices <- str_sub(files, 52, -15)

# County results

statewide_county <- data.frame(county = character(),
                               office = character(),
                               candidate = character(),
                               votes = numeric())

for (y in offices) {
  index <- which(offices == y)
  assign(y, read_excel(files[index], skip = 5) %>%
           slice(1:100) %>%
           filter(!(str_detect(County, "District \\b\\b"))) %>%
           mutate(office = y) %>%
            gather(key = candidate, value = votes, -County, -office))
  statewide_county <- rbind.data.frame(statewide_county, get(y)) %>%
    arrange(office, County)
}

write.csv(statewide_county, file = "2016/20161108__nd__general__county.csv")

# Precinct results

counties <- unique(statewide_county$County)

