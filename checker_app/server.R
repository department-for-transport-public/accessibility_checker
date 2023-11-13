
library(shiny)
library(tidyods)
library(dplyr)
library(tidyr)
library(DT)
library(purrr)
library(stringr)
library(xml2)

##Load in R functions
source("functions.R")
##Source everything in the modules folder
list.files("modules", full.names = TRUE) %>%
  purrr::walk(.f = source)

# Define server logic required to draw a histogram
shinyServer(function(input, output, session) {

  xlsx_Server("xlsx")
  ods_Server("ods")

})
