#Function to produce nicely formatted DT table

datatbl <- function(data){ 

DT::datatable(
  data,
  rownames = FALSE,
  options = list(sDom  = '<"top">rt',
                 columnDefs = list(list(targets = ncol(data) - 1, 
                                        visible = FALSE)),
                 pageLength = 50)) %>% 
    formatStyle(
    "check",
    target = 'row',
    backgroundColor = styleEqual(c(FALSE, TRUE), c('orange', 'lightgray')))
}


##Data table output with a spinner
spinDT <- function(table){
  
  DT::dataTableOutput(table) %>%
    shinycssloaders::withSpinner()
}

##Sequence along data worksheets
seq_along_wb <- function(worksheets){
  
  sheets <- c()
  
  ##Check which worksheets are data
  for(i in seq_along(worksheets))
    if("WorkSheet" %in% class(worksheets[[i]])){
      sheets <- c(sheets, i)
    }
  
  return(sheets)
}


##Extract a specific xml content from an ODS file
extract_xml <- function(file, content_type = "content.xml"){
  #create a temp folder
  temp_dir <- tempdir()
  ##Unzip file into it
  unzip(file, exdir = temp_dir, files = content_type)
  
  # Read and parse the xml file
  xml2::read_xml(file.path(temp_dir, content_type))
}

##Util function to return nulls as empty string
null_s <- function(x){
  
  if(is.null(x)){
    x <- ""
  }
  
  if(length(x) == 0){
    x <- ""
  }
  
  return(x)
}