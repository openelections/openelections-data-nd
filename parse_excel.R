# Parse ND .xls files

library(readxl)
library(dplyr)
library(stringr)
library(readr)
library(tidyr)
library(lettercase)

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
           mutate(Office = str_title_case(y)) %>%
            gather(key = Candidate, value = Votes, -County, -Office))
  statewide_county <- rbind.data.frame(statewide_county, get(y)) %>%
    arrange(Office, County)
}

write_csv(statewide_county, "2016/20161108__nd__general__county.csv")

# Precinct results

counties <- unique(statewide_county$County)

statewide_precinct <- data.frame(Precinct = character(),
                                 Office = character(),
                                 County = character(),
                                 Candidate = character(),
                                 Votes = character())
                               
for (y in offices) {
  index <- which(offices == y)
  for(z in counties) {
    assign(y, read_excel(files[index], sheet = z, skip = 5) %>%
             filter(!(str_detect(Precinct, "Results") | 
                        is.na(Precinct) |
                        Precinct == "TOTAL")) %>%
             mutate(Office = str_title_case(y), County = z) %>%
             gather(key = Candidate, value = Votes, -County, -Precinct, -Office))
    statewide_precinct <- rbind.data.frame(statewide_precinct, get(y))
  }
}

statewide_precinct <- statewide_precinct %>%
  select(County, everything()) %>%
  mutate(Votes = as.numeric(str_extract(Votes, "^\\d*"))) %>%
  arrange(Office, County, Precinct)

write_csv(statewide_precinct, "2016/20161108__nd__general__precinct.csv")
