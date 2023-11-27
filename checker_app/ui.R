library(shiny)
library(DT)
library(shinythemes)
library(purrr)

##Load in R functions
source("functions.R")

##Source everything in the modules folder
list.files("modules", full.names = TRUE) %>%
  purrr::walk(.f = source)

# Define UI 
shinyUI(
  fluidPage(
  
  ##Making our tabs green
  tags$head(
    tags$link(rel = "stylesheet", type = "text/css", href = "mytheme.css")
  ),
  
    navbarPage(collapsible = TRUE,
      title = "Spreadsheet Accessibility Checker",
      tabPanel("ODS",
               ods_UI("ods", "ods")),#End of tab panel
      tabPanel("XLSX",
               xlsx_UI("xlsx", "xlsx")
      ), #End of tab panel
       tabPanel("Instructions",
                h2("How to use"),
                
                "There are two tabs in the checker, depending on whether you're checking xlsx or ods files. Please select the tab relevant to the file type you're uploading.",
                br(),
                "You can upload one or more files at a time, the checker will check each of them in turn. Large files may be slow to check, please be patient while the results load!",
                br(),  
                "Click on each tab in turn to see the results of the checks. Significant results are highlighted in orange, all other content is provided for your information only.",
                br(),
                "Some content cannot be checked automatically at this stage, and you will still need to perform manual checks to ensure these aspects of accessibility are met. You can find these on the", em("manual checks"), "tab.",
                br(),
                "Please note that in many cases, more detailed checks can be carried out on ODS format in comparison to XLSX files, but as a result the ODS checker is slower to load.",
                
                br(),
                h2("Data security"),
                
                "This checker does not save any of the data uploaded to it, and is hosted on an EU-based server. It is suitable for use for any published statistics or for Official statistics prior to publication,", strong("but should not be used for data beyond OFFSEN designation")
       )
    ) #end of navbar panel
  ) #End of fluid page
)#End of UI
