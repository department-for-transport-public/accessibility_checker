xlsx_UI <- function(id, label) {
  ns <- NS(id)

  #Start the bulk of the UI here------------------
  tagList(
    # Sidebar with a slider input for number of bins
    sidebarLayout(
      sidebarPanel(
        shiny::fileInput(inputId = ns("excel_files"),
                         label = h3("Please upload your tables here"),
                         accept = ".xlsx", 
                         multiple = TRUE),
        hr(),
        actionButton(ns("show"), 
                     "Need help?"),
        
        width = 3
      ),#end of tab panel,
      
      # Show a plot of the generated distribution
      mainPanel(
        
        tabsetPanel(type = "pills",
                    
                    tabPanel("Spreadsheet",
                             h2("Spreadsheet contains a cover worksheet"),
                             spinDT(ns("cover_sheet")),
                             br(),
                             h2("There should be no blank sheets"),
                             spinDT(ns("empty_sheets")),
                             br(),
                             h2("Spreadsheet contains a table of contents in a separate sheet"),
                             spinDT(ns("contents_sheet")),
                             br(),
                             h2("Spreadsheet contains a Notes worksheet"),
                             spinDT(ns("notes_sheet")),
                             br(),
                             h3("(Recommended) A1 should be the active cell"),
                             spinDT(ns("active_cell")),
                             br(),
                             h3("(Recommended) The workbook should have a title"),
                             spinDT(ns("book_title")),
                             br()
                             ), #End of panel
                    
                    tabPanel("Worksheets",
                             h2("Worksheet does not include images"),
                             spinDT(ns("no_images")),
                             br(),
                             h2("Links should have a descriptive name"),
                             spinDT(ns("descriptive_links")),
                             br(),
                             h3("(Recommended) There is only one table per worksheet"),
                             spinDT(ns("single_table")),
                             br(),
                             h3("(Recommended) There should be no frozen panes"),
                             spinDT(ns("frozen_panes"))
                             ),#End of tab panel
                    
                    
                    tabPanel("Tables",
                             h2("Tables should start in column A"),
                             spinDT(ns("col_a")),
                             br(),
                             h2("All content should be in a marked-up table"),
                             spinDT(ns("no_table")),
                             br(),
                             h2("All tables should have meaningful names"),
                             spinDT(ns("named_tables")),
                             br(),
                             h2("There should be no merged cells"),
                             spinDT(ns("merged_cells")),
                             br(),
                             h2("There should be no empty rows or columns"),
                             spinDT(ns("empty_cols")),
                             br(),
                             h3("(Recommended) There should be no hidden columns or rows"),
                             spinDT(ns("hidden_cols")),
                             br(),
                             h3(" (Recommended) Tables should have filter buttons turned off"),
                             spinDT(ns("hidden_filter")),
                             br()
                             
                    ),  #End of panel
                    
                    tabPanel("Formatting",
                             
                             "Formatting checks not available for xlsx format tables.
                             Please try the ODS checker instead."
                             
                    ), #End of panel
                    
                    tabPanel("Table content",
                             h2("Do not use superscript to signpost to notes"),
                             spinDT(ns("superscript"))
                             ), #End of panel
                    
                    tabPanel("Manual checks",
                             h2("The following mandatory checks cannot yet be done in this tool. 
                                Please ensure your table meets the following before publication:"),
                             br(),
                             tags$ul(
                               tags$li("Worksheet has unique tab names"),
                               tags$li("A description of the worksheet is provided in cell A2 or A3"),
                               tags$li("Cover sheet includes title of spreadsheet in cell A1"),
                               tags$li("There are no spelling or grammar errors"),
                               tags$li("Acronyms and abbreviations are expanded when first used")
                             )
                    )#End of tab panel
        ) #End of tabset panel
        
      )#End of main panel
    ) #End of sidebar layout
  )#End of tag list
}# End of UI


##Server-------------

xlsx_Server <- function(id) {
  moduleServer(
    id,
    function(input, output, session) {
      ns <- NS(id)
      
    
      ##Read in the file as a workbook 
      workbook_format <- reactive({
        
        validate(
          need(input$excel_files != "", "Please select a file to check")
        )
        
        ##Loop over all loaded files
        wb_list <- purrr::map(.x = input$excel_files$datapath,
                              .f = openxlsx::loadWorkbook)
        
        ##Create a named list
        names(wb_list) <- input$excel_files$name
        return(wb_list)
        
      })
      
      ##Read sheet names from workbook 
      workbook_sheet_names <- reactive({
        
        validate(
          need(input$excel_files != "", "Please select a file to check")
        )
        
        ##Loop over all loaded files
        wb_list <- purrr::map(.x = input$excel_files$datapath,
                              .f = openxlsx::getSheetNames)
        
        ##Create a named list
        names(wb_list) <- input$excel_files$name
        return(wb_list)
        
      })
      
      #List sheets with containing table markup on them 
      tables_and_sheets <- reactive({
        
        no_table_sheets <- function(sheet_names, wb) {
          
          tables <- openxlsx::getTables(wb = wb, sheet = sheet_names)
          
          tibble(
            "sheet" = sheet_names,
            "range" = if(length(tables) != 0){names(tables)} else{""},
            "names" = if(length(tables) != 0){tables} else{""})
          
        }
        
        ##Function to loop that over all sheets in a table
        loop_sheets <- function(wb, names){
          purrr::map_df(.x = names,
                        .f = no_table_sheets,
                        wb = wb)
        }
        
        # ##Loop over all tables
        purrr::map2_df(.x = workbook_format(),
                       .y = workbook_sheet_names(),
                       .f = loop_sheets,
                       .id = "table") 
        
      })
      
      ##Create output tables----------------------------------------------------
      
      ##There should be a cover sheet-------------
      output$cover_sheet <- DT::renderDataTable({
        
        ##Create a table of sheet names and workbooks
        tibble::tibble("sheet" = workbook_sheet_names(),
                       "workbook" = names(workbook_sheet_names())) %>%
          ##Check if there's a cover sheet
          dplyr::mutate(check = grepl("cover", sheet, ignore.case = TRUE)) %>%
          
          datatbl()
      })
      
      
      #There should be no empty worksheets---------------------------
      
      output$empty_sheets <- DT::renderDataTable({
        
        #Function to check if a sheet is empty
        check_empty <- function(i, wb){
          tibble(
            "sheet" = wb$sheet_names[[i]],
            "data_objects" = wb$worksheets[[i]]$field(name = "sheet_data")$field(name = "n_elements") 
          )
        }
        
        ##Map it over every sheet in an object
        check_empty_sheets <- function(wb){
          purrr::map_df(.x = seq_along_wb(wb$worksheets),
                        .f = check_empty,
                        wb = wb)
        }
        
        #Map it over every workbook
        purrr::map_df(.x = workbook_format(),
                      .f = check_empty_sheets,
                      .id = "workbook") %>%
          dplyr::mutate(check = data_objects != 0) %>%
          datatbl()
        
        
      })
      
      
      ##There should be a table of contents-------------
      output$contents_sheet <- DT::renderDataTable({
        
        ##Create a table of sheet names and workbooks
        tibble::tibble("sheet" = workbook_sheet_names(),
                       "workbook" = names(workbook_sheet_names())) %>%
          ##Check if there's a contents sheet
          dplyr::mutate(check = grepl("content", sheet, ignore.case = TRUE)) %>%
          
          datatbl()
      })
      
      ##There should be a notes sheet-------------
      output$notes_sheet <- DT::renderDataTable({
        
        ##Create a table of sheet names and workbooks
        tibble::tibble("sheet" = workbook_sheet_names(),
                       "workbook" = names(workbook_sheet_names())) %>%
          ##Check if there's a notes sheet
          dplyr::mutate(check = grepl("note", sheet, ignore.case = TRUE)) %>%
          
          datatbl()
      })
      
      ##The active cell in each sheet should be A1-----------------
      output$active_cell <- DT::renderDataTable({
        
        #Function to check what cells are active
        check_active <- function(i, wb){
          tibble(
            "sheet" = wb$sheet_names[[i]],
            "active_cell" = stringr::str_extract(
              pattern = 'sqref[=]["]([A-Z]{1,2}\\d{1,2})', 
              string = wb$worksheets[[i]]$field(name = "sheetViews"),
              group = 1) 
            
          )
        }
        
        ##Map it over every sheet in an object
        check_active_cells <- function(wb){
          purrr::map_df(.x = seq_along_wb(wb$worksheets),
                        .f = check_active,
                        wb = wb)
        }
        
        #Map it over every workbook
        purrr::map_df(.x = workbook_format(),
                      .f = check_active_cells,
                      .id = "workbook") %>%
          na.omit() %>%
          dplyr::mutate(check = active_cell == "A1") %>%
          datatbl()
        
        
      })
      
      
      #The workbook should have a title--------------------
      output$book_title <- DT::renderDataTable({
        
        ##Function to Look for workbook title
        wb_title <- function(wb) {
          if(grepl("[<]dc\\:title\\>", wb$core)){
            book_title <- gsub(".*[<]dc\\:title\\>(.*)\\<*dc\\:title\\>.*", "\\1", wb$core)
          } else{
            book_title <- ""
          }
        }
        
        ##Loop activeSheet reading over objects
        purrr::map_dfr(.x = workbook_format(),
                       .f = wb_title) %>%
          tidyr::pivot_longer(cols = dplyr::everything(),
                              names_to = "Workbook",
                              values_to = "Workbook title") %>%
          dplyr::mutate(check = `Workbook title` != "") %>%
          
          datatbl()
        
        
      })
      
      
      ###Check there are no images in the workbook--------------
      
      output$no_images <- DT::renderDataTable({
        
        ##Function to Look for workbook title
        wb_img <- function(wb) {
          if(length(wb$media) > 0){
            number_img <- length(wb$media)
          } else{
            number_img <- 0
          }
        }
        
        ##Loop activeSheet reading over objects
        purrr::map_dfr(.x = workbook_format(),
                       .f = wb_img) %>%
          tidyr::pivot_longer(cols = dplyr::everything(),
                              names_to = "Workbook",
                              values_to = "Number of images") %>%
          dplyr::mutate(check = `Number of images` == 0) %>%
          
          datatbl()

      })
      
      
      #Make hyperlink text descriptive--------------------
      
      output$descriptive_links <- DT::renderDataTable({
        
        #Function to check whether link description is descriptive
        check_links <- function(i, wb){
          
          t <- tibble()
          
          for(j in seq_along(wb$worksheets[[i]]$field(name = "hyperlinks"))){
            
            wb$worksheets[[i]]$field(name = "hyperlinks")[[j]]$field(name = "target")
            
            ref <- wb$worksheets[[i]]$field(name = "hyperlinks")[[j]]$field(name = "ref")
            
            content <- read.xlsx(wb, 
                                 sheet = i,
                                 rows = gsub("[A-Z]", "", ref), 
                                 cols = openxlsx::convertFromExcelRef(gsub("[0-9]", "", ref)), 
                                 colNames = FALSE)
            
            t <-  bind_rows( t,
                             tibble(
                               "sheet" = wb$sheet_names[[i]],
                               "cell_link" = ref,
                               "link_name" = content[[1, 1]]
                             ) )
            
          }
          
          return(t)
          
        }
        
        ##Map it over every sheet in an object
        check_all_links <- function(wb){
          purrr::map_df(.x = seq_along_wb(wb$worksheets),
                        .f = check_links,
                        wb = wb)
        }
        
        #Map it over every workbook
        purrr::map_df(.x = workbook_format(),
                      .f = check_all_links,
                      .id = "workbook") %>%
          dplyr::mutate(check = !grepl("^http", link_name) &  !grepl("link|here", link_name)) %>%
          datatbl()
        
        
      })
      
      ##The workbook should have one table per sheet-------------
      output$single_table <- DT::renderDataTable({
        
        tables_and_sheets() %>%
          ##Count there's at least one table per sheet
          dplyr::group_by(table, sheet) %>%
          dplyr::summarise(count = sum(names != "")) %>%
          dplyr::ungroup() %>%
          dplyr::mutate(check = count > 1) %>%
          datatbl()
        
      })
      
      
      #There should be no frozen panes---------------------------
      
      output$frozen_panes <- DT::renderDataTable({
        
        #Function to check if a sheet has hidden cols
        check_frozen <- function(i, wb){
          tibble(
            "sheet" = wb$sheet_names[[i]],
            "frozen_pane" = length(wb$worksheets[[i]]$field(name = "freezePane")) != 0
          )
        }
        
        ##Map it over every sheet in an object
        check_frozen_pane <- function(wb){
          purrr::map_df(.x = seq_along_wb(wb$worksheets),
                        .f = check_frozen,
                        wb = wb)
        }
        
        #Map it over every workbook
        purrr::map_df(.x = workbook_format(),
                      .f = check_frozen_pane,
                      .id = "workbook") %>%
          dplyr::mutate(check = as.numeric(frozen_pane == FALSE)) %>%
          datatbl()
        
        
      })
      
      ##Tables should start in column A-------------------
      
      output$col_a <- DT::renderDataTable({
        
        tables_and_sheets() %>%
          ##Drop empty ranges 
          dplyr::filter(range != "") %>%
          ##Get the first number from the range
          dplyr::mutate(row = gsub("^[:A-Z:](.*)[:].*", "\\1", range)) %>%
          ##Group by sheet and find the top table
          dplyr::group_by(table, sheet) %>%
          #Check if ranges start from A
          dplyr::filter(range == max(range, na.rm = TRUE)) %>%
          dplyr::ungroup() %>%
          ##Check that the first column is A
          dplyr::mutate(check = grepl("^A\\d", range))%>%
          datatbl()
        
      })
      
      ##The workbook should have tables on every sheet-------------
      output$no_table <- DT::renderDataTable({
        
        tables_and_sheets() %>%
          ##Count there's at least one table per sheet
          dplyr::group_by(table, sheet) %>%
          dplyr::summarise(count = sum(names != "")) %>%
          dplyr::ungroup() %>%
          dplyr::mutate(check = count > 0) %>%
          datatbl()
        
      })
      
      ##All tables should have a meaningful name--------------------
      output$named_tables <- DT::renderDataTable({
        
        tables_and_sheets() %>%
          dplyr::filter(range != "") %>%
          dplyr::mutate(check = (names != "" & !grepl("^Table\\d.*", names))) %>%
          datatbl()
        
      })
      
      
      #No merged cells--------------------
      
      output$merged_cells <- DT::renderDataTable({
        
        #Function to check what rows and columns are merged
        check_merged <- function(i, wb){
          
          merged <- stringr::str_extract(pattern = '\\"([A-Z].*\\d)\\"',
                                         string = wb$worksheets[[i]]$field(name = "mergeCells"),
                                         group = 1)
          
          tibble(
            "sheet" = wb$sheet_names[[i]],
            "merged_cells" = if(length(merged) == 0){""} else{merged}
          ) 
          
        }
        
        ##Map it over every sheet in an object
        check_merged_all <- function(wb){
          purrr::map_df(.x = seq_along_wb(wb$worksheets),
                        .f = check_merged,
                        wb = wb)
        }
        
        #Map it over every workbook
        purrr::map_df(.x = workbook_format(),
                      .f = check_merged_all,
                      .id = "workbook") %>%
          dplyr::mutate(check = merged_cells == "") %>%
          datatbl()
        
        
      })
      
      
      #No empty rows and columns--------------------
      
      output$empty_cols <- DT::renderDataTable({
        
        #Function to check what rows and columns are missing
        check_missing <- function(i, wb){
          
          rows <- wb$worksheets[[i]]$field(name = "sheet_data")$field(name = "rows")
          cols <- wb$worksheets[[i]]$field(name = "sheet_data")$field(name = "cols")
          
          ##Get rid of NA stuff
          rows <- c(rows[!is.na(wb$worksheets[[i]]$field(name = "sheet_data")$field("t"))],
                    rows[!is.na(wb$worksheets[[i]]$field(name = "sheet_data")$field("v"))],
                    rows[!is.na(wb$worksheets[[i]]$field(name = "sheet_data")$field("f"))])
          
          cols <- c(cols[!is.na(wb$worksheets[[i]]$field(name = "sheet_data")$field("t"))],
                    cols[!is.na(wb$worksheets[[i]]$field(name = "sheet_data")$field("v"))],
                    cols[!is.na(wb$worksheets[[i]]$field(name = "sheet_data")$field("f"))])
          
          tibble(
            "sheet" = wb$sheet_names[[i]],
            "missing_rows" = paste(setdiff(1:max(rows), unique(rows)), collapse = ", "),
            "missing_cols" = paste(setdiff(1:max(cols), unique(cols)), collapse = ", ")
          ) 
          
        }
        
        ##Map it over every sheet in an object
        check_missing_all <- function(wb){
          purrr::map_df(.x = seq_along_wb(wb$worksheets),
                        .f = check_missing,
                        wb = wb)
        }
        
        #Map it over every workbook
        purrr::map_df(.x = workbook_format(),
                      .f = check_missing_all,
                      .id = "workbook") %>%
          dplyr::mutate(check = (missing_rows == "" & missing_cols == "")) %>%
          datatbl()
        
        
      })
      
      #There should be no hidden columns---------------------------
      
      output$hidden_cols <- DT::renderDataTable({
        
        #Function to check if a sheet has hidden cols
        check_hidden <- function(i, wb){
          tibble(
            "sheet" = wb$sheet_names[[i]],
            "hidden_cols" = sum(as.numeric(attr(wb$colWidths[[i]], "hidden")))
          )
        }
        
        ##Map it over every sheet in an object
        check_hidden_cols <- function(wb){
          purrr::map_df(.x = seq_along_wb(wb$worksheets),
                        .f = check_hidden,
                        wb = wb)
        }
        
        #Map it over every workbook
        purrr::map_df(.x = workbook_format(),
                      .f = check_hidden_cols,
                      .id = "workbook") %>%
          dplyr::mutate(check = hidden_cols == 0) %>%
          datatbl()
        
        
      })
      
      #Tables should have filter buttons switched off---------------------
      
      output$hidden_filter <- renderDataTable({
        
        ##Function to check for filters in place
        check_filters <- function(wb){
          tibble(
            "names" = attr(wb$tables, which = "tableName"),
            "sheet_number" = attr(wb$tables, "sheet"),
            "filter_hidden" = grepl("hiddenButton", wb$tables)
            
          )
        }
        
        ##Loop check over all sheets
        purrr::map_df(.x = workbook_format(),
                      .f = check_filters,
                      .id = "workbook") %>%
          ##Check buttons have been removed
          dplyr::mutate(check = as.numeric(filter_hidden))%>%
          datatbl()
      })
      
      ##Cells that contain superscript------------------------------------
      
      output$superscript <- DT::renderDataTable({
        
        #Function to check what rows and columns are merged
        check_super <- function(wb){
          
          #Find all instances of superscript in workbook strings
          text <- gsub("^.*[>]([^<].*?[^>])[<].*", "\\1", 
                       wb$sharedStrings[grepl("superscript", wb$sharedStrings)], perl = TRUE)
          
          #Keep anything that contains no <> symbols but text
          
          tibble(
            "superscript_string" = text[!grepl("[<]|[>]", text) & grepl("[A-Za-z]", text)]
          ) 
          
        }
        
        
        #Map it over every workbook
        purrr::map_df(.x = workbook_format(),
                      .f = check_super,
                      .id = "workbook") %>%
          dplyr::mutate(check = 0) %>%
          datatbl()
        
        
      })
      
      #The first worksheet should be open when the workbook is saved----------
      output$first_sheet <- DT::renderDataTable({
        
        ##Loop activeSheet reading over objects
        purrr::map_dfr(.x = workbook_format(),
                       .f = openxlsx::activeSheet) %>%
          tidyr::pivot_longer(cols = dplyr::everything(),
                              names_to = "Workbook",
                              values_to = "Open sheet") %>%
          dplyr::mutate(check = `Open sheet` == 1) %>%
          
          datatbl()
        
        
      })
      
      
     ##There should be nothing under a table--------------------------
      
      output$under_table <- DT::renderDataTable({
        
        #Function to check what rows are filled under the table
        check_under_table <- function(i, wb){
          
          ##Find last table rows and cols
          table <- names(attr(wb$worksheets[[i]]$field(name = "tableParts"), "tableName"))
          last_row <- gsub(".*[:][A-Z]{1,2}(.*)$", "\\1", table)
          
          rows <- wb$worksheets[[i]]$field(name = "sheet_data")$field(name = "rows")
          
          ##Get rid of NA stuff
          rows <- c(rows[!is.na(wb$worksheets[[i]]$field(name = "sheet_data")$field("t"))],
                    rows[!is.na(wb$worksheets[[i]]$field(name = "sheet_data")$field("v"))],
                    rows[!is.na(wb$worksheets[[i]]$field(name = "sheet_data")$field("f"))])
          
          ##If last_row is null, just swap it for the max rows value
          if(length(last_row) == 0){
            last_row <- max(rows)
          }
          
          tibble(
            "sheet" = wb$sheet_names[[i]],
            "rows_under_table" = paste(setdiff(last_row:max(rows), unique(rows)), collapse = ", ")
          ) 
          
        }
        
        ##Map it over every sheet in an object
        check_under_table_all <- function(wb){
          purrr::map_df(.x = seq_along_wb(wb$worksheets),
                        .f = check_under_table,
                        wb = wb)
        }
        
        #Map it over every workbook
        purrr::map_df(.x = workbook_format(),
                      .f = check_under_table_all,
                      .id = "workbook") %>%
          dplyr::mutate(check = rows_under_table == "" ) %>%
          datatbl()
        
        
      })
     
      
      
      ##There should be no header/footer----------------------------
      
      output$header_footer <- DT::renderDataTable({
        
        #Function to check if a sheet has hidden cols
        check_head_foot <- function(i, wb){
          
          if(length(wb$worksheets[[i]]$headerFooter) != 0){
            
            header <- wb$worksheets[[i]]$headerFooter$oddHeader[[2]]
            footer <- wb$worksheets[[i]]$headerFooter$oddFooter[[2]]
          } else{
            header <- ""
            footer <- ""
          }
          
          tibble(
            "sheet" = wb$sheet_names[[i]],
            "header" = header,
            "footer" = footer
          )
        }
        
        ##Map it over every sheet in an object
        check_head_foot_all <- function(wb){
          purrr::map_df(.x = seq_along_wb(wb$worksheets),
                        .f = check_head_foot,
                        wb = wb)
        }
        
        #Map it over every workbook
        purrr::map_df(.x = workbook_format(),
                      .f = check_head_foot_all,
                      .id = "workbook") %>%
          dplyr::mutate(check = as.numeric(header == "" & footer == "")) %>%
          datatbl()
        
        
      })
      
      
      
      ##Help message popup-----------------------------
      observeEvent(input$show, {
        showModal(modalDialog(
          title = "How to use the accessibility checker",
          "Click the 'Browse' button to upload one or more publication tables
      in .xlsx format into the file input. 
      Click through the different tabs to see the results of the various check types.
      The file checker will return highlighted orange rows when it finds potential accessibility
      issues.",
      easyClose = TRUE,
      footer = NULL
        ))
      })
   
    }   
  )
}