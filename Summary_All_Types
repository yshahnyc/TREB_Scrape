#load Libraries
library(tidyverse)
library(tabulizer)
library(RCurl)
library(shiny)
library(miniUI)
library(openxlsx)
library(lubridate)

#date prep - develop a list of dates (yymm) to push to the list of pdf documents to pull in the url
datstart <- 9601
datend <- as.numeric(format(Sys.Date(), "%y%m"))-1

datseq <- function(t1, t2) {
  format(seq(as.Date(paste0(t1,"01"), "%y%m%d"), 
             as.Date(paste0(t2,"01"), "%y%m%d"),
             by="month"),
         "%y%m")
         }

mwseq <- datseq(datstart, datend)

mwseq <- append(mwseq,
                as.numeric(format(Sys.Date(), "%y%m"))-1,
                after = length(mwseq))

rm(datend,datstart, datseq)


#make a list of the file links
for (i in 1:length(mwseq)) {
  pdf_list <- paste0("http://trreb.ca/files/market-stats/market-watch/mw",mwseq,".pdf")
}

#define tab areas for the tables you're looking to pull (does a better job than Tabulizer's default table pull)
pdfcontent <- "http://trreb.ca/files/market-stats/market-watch/mw2009.pdf"
tab_areas <- locate_areas(pdfcontent,
                          pages = 3,
                          widget = "shiny")

#Summary data for time period of choice.
pdf_list_select <- pdf_list[265:276]
mwseq_select <- mwseq[265:276]

#list of table names
for (i in 1:length(mwseq_select)) {
  table_name <- paste0("summary_alltypes", mwseq_select)
}

#create a dummy list and extract data
pdf_results <- list()

for (i in 1:length(pdf_list_select)) {
  extract_tables(pdf_list_select[i],
                 pages = 3,
                 area = tab_areas,
                 guess = FALSE,
                 method = "lattice",
                 output = "data.frame",
                 header = TRUE) ->
    pdf_results[[i]]
}

names(pdf_results) <- table_name
pdf_results$columns <- colnames(pdf_results$summary_alltypes1801[[1]])

#export to excel
write.xlsx(pdf_results,
           "outputs/pdf_results_18.xlsx",
           col.names = TRUE)
