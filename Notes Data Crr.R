
## Enable Trust access to the VBA project object model
library(rvest)
library(stringr)
library(readxl)
library(dplyr)
library(reshape2)
library(lubridate)
library(ggplot2)
library(scales)

getOption("max.print")
options(max.print=1000000)


## Enable Trust access to the VBA project object model
shell(shQuote(normalizePath("C:/Users/user.userov/Documents/Notes/current/VBS1.vbs")), "cscript", flag = "//nologo")

getwd()
setwd("C:/Users/user.userov/Documents/Notes/current")

file.list <- list.files(pattern='*.xls') ## get all the excel files in the directory
df.list <- lapply(file.list, read_excel) ## read them all

dat = Reduce(function(x, y) merge(x, y, by = "Code",
                                  all.x = TRUE, all.y = TRUE), df.list)
dt = dat[ , order(colnames(dat))]
Code = dt$Code

dt = dt[, -ncol(dt)]  ## Code column is at the end. Delete it 
dt = cbind(Code, dt)  ## and move it forward

dt[ , "URL"] <- NA   ## create new NA URL columns
dt$URL = apply(dt, 1, function(x) paste0("http://old.kase.kz/en/gsecs/show/", toString(x[1])))

name = as.character(dt$Code)
URL = dt$URL

instruments <- data.frame(name, URL, stringsAsFactors = FALSE)
instruments   ## ALWAYS check NTK091_1340 and NTK092_1394
instruments = instruments[-c(498, 559), ]  ## find and delete. no data for them

wanted <- c("NSIN" = "ISIN",
            "Nominal value in issue's currency" = "NominalValue",
            "Number of bonds outstanding" = "BondsQuantity",
            "Issue volume, KZT" = "IssueVolume",
            "Date of circulation start" = "StartDate",
            "Principal repayment date" = "EndDate")

getValues <- function (name, url) {
  df <- url %>%
    read_html() %>%
    html_nodes("table.top") %>%
    html_table()
  df = as.data.frame(df)
  names(df) <- c("full_name", "value")
  
  
  ## filter and remap wanted columns
  result <- df[df$full_name %in% names(wanted),]
  result$column_name <- sapply(result$full_name, function(x) {wanted[[x]]})
  
  ## add the identifier to every row
  result$name <- name
  return (result[,c("name", "column_name", "value")])
}

## invoke function for each name/URL pair - returns list of data frames
columns <- apply(instruments[,c("name", "URL")], 1, function(x) {getValues(x[["name"]], x[["URL"]])})

## bind using dplyr:bind_rows to make a tall data frame
tall <- bind_rows(columns)

## make wide using dcast from reshape2
wide <- dcast(tall, name ~ column_name, id.vars = "value")
colnames(wide)[1]= "Code"

## ***********************************************************
## ***********************************************************

dim(dt)
dim(wide)
dt = dt[-c(498, 559), ]  ## UPDATE this lime accordingly


nwd = merge(dt, wide, by = "Code")
nwd

write.csv(nwd, "nwdcurrent.csv")
