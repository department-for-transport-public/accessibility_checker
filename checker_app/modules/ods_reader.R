ods_UI <- function(id, label) {
  ns <- NS(id)
  
  #Start the bulk of the UI here------------------
  tagList(
    # Sidebar with a slider input for number of bins
    sidebarLayout(
      sidebarPanel(
        shiny::fileInput(inputId = ns("ods_files"),
                         label = h3("Please upload your tables here"),
                         accept = ".ods", 
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
                             h4("This will always be the case in an ODS file"),
                             h3("(Recommended) The workbook should have a title"),
                             spinDT(ns("book_title"))
                    ), #End of panel
                    
                    tabPanel("Worksheets",
                             h2("Worksheet does not include images"),
                             spinDT(ns("no_images")),
                             br(),
                             h2("Links should have a descriptive name"),
                             spinDT(ns("named_links")),
                             br(),
                             h3("(Recommended) There is only one table per worksheet"),
                             spinDT(ns("single_table")),
                             br(),
                             h3("(Recommended) There should be no frozen panes"),
                             h4("This will always be the case in an ODS file")
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
                             h2("All text is set to sans serif font (e.g. Arial)"),
                             spinDT(ns("font_family")),
                             h2("All text is wrapped and visible"),
                             spinDT(ns("wrapped")),
                             h2("Table title is set to 'Heading 1'"),
                             spinDT(ns("headings")),
                             h3("(Recommended) Font is at least size 10"),
                             spinDT(ns("font_size")),
                             h3("(Recommended) Font colour is 'automatic'"),
                             spinDT(ns("font_colour")),
                             h3("(Recommended) No background fill"),
                             spinDT(ns("background_fill")),
                             h3("(Recommended) Text is horizontal"),
                             spinDT(ns("text_horizontal")),
                             h3("(Recommended) Italics or underlining are not used for emphasis"),
                             spinDT(ns("italic_underline"))
                             
                    ), #End of panel
                    
                    tabPanel("Table content",
                             h2("Do not use superscript to signpost to notes"),
                             spinDT(ns("superscript")),
                             br(),
                             h3("(Recommended) Dates are written in order: date, month, year."),
                             spinDT(ns("date_style")),
                             h3("(Recommended) Thousands are separated by commas"),
                             spinDT(ns("comma_format"))
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

ods_Server <- function(id) {
  moduleServer(
    id,
    function(input, output, session) {
      ns <- NS(id)
      
      #source those functions 
      source("functions.R")
      
      
      ##Read in the file in the tidyods format 
      tidy_format <- reactive({
        
        validate(
          need(input$ods_files != "", "Please select a file to check")
        )
        
        ##Function to read all sheets
        read_all_sheets <- function(x){
          
          sheets <- tidyods::ods_sheets(x)
          names(sheets) <- sheets
          
          purrr::map_df(.x = sheets,
                        .f = tidyods::read_ods_cells,
                        path = x,
                        quiet = TRUE,
                        .id = "sheet")
        }
        
        files <- input$ods_files$datapath
        names(files) <- input$ods_files$name
        
        withProgress(message = 'Loading data', value = 0,{
        
          
        ##Loop over all loaded files
        wb_list <- purrr::map(.x = files,
                              .f = function(x){
                                incProgress(1/length(files)) 
                                read_all_sheets(x)
                              }
                                ) %>%
          data.table::rbindlist(idcol = "Workbook") %>%
          #drop weird sheet names
          dplyr::filter(!grepl("file://", sheet, fixed = TRUE))
        
        return(wb_list)
        
        
        })
        
      })
      
      ##Read sheet names from workbook 
      workbook_sheet_names <- reactive({
        
        validate(
          need(input$ods_files != "", "Please select a file to check")
        )
        
        files <- input$ods_files$datapath
        names(files) <- input$ods_files$name
        
        ##Loop over all loaded files
        purrr::map_df(.x = files,
                      .f = function(x){
                        tibble(
                          "sheet" = tidyods::ods_sheets(x))},
                      .id = "Workbook") 
        
        
      })
      
      ##Raw content sheets for each book
      book_content <- reactive({
        
        validate(
          need(input$ods_files != "", "Please select a file to check")
        )
        
        files <- input$ods_files$datapath
        names(files) <- input$ods_files$name
        
        ##List of all content sheets
        purrr::map(.x = files,
                   .f = extract_xml)
        
      })
      
      ##Raw style sheets for each book
      book_style <- reactive({
        
        validate(
          need(input$ods_files != "", "Please select a file to check")
        )
        
        files <- input$ods_files$datapath
        names(files) <- input$ods_files$name
        
        ##List of all content sheets
        purrr::map(.x = files,
                   .f = extract_xml,
                   content_type = "styles.xml")
        
      })
      
      ##Raw metadata sheets for each book
      book_meta <- reactive({
        
        validate(
          need(input$ods_files != "", "Please select a file to check")
        )
        
        files <- input$ods_files$datapath
        names(files) <- input$ods_files$name
        
        ##List of all content sheets
        purrr::map(.x = files,
                   .f = extract_xml,
                   content_type = "meta.xml")
        
      })
      
      ##Table content for each workbook
      book_tables <- reactive({
        
        table_data <- function(xml){
          table <- xml2::xml_find_all(xml, "//table:database-range")
          
          tibble(
            "table_name" = table %>%
              xml2::xml_attr("name") %>%
              null_s(),
            "cell_range" = table %>%
              xml2::xml_attr("target-range-address") %>%
              null_s(),
            "filter_buttons" =  table %>%
              xml2::xml_attr("display-filter-buttons") %>%
              null_s()) %>%
            ##Get names and cell ranges separate
            tidyr::separate(col = cell_range, into = c("start", "end"), sep = "[:]") %>%
            tidyr::separate(start, into = c("sheet", "start_cell"), sep = "[.]") %>%
            tidyr::separate(end, into = c("sheet", "end_cell"), sep = "[.]") 
        }
        
        
        purrr::map_df(.x = book_content(),
                      .f = table_data, 
                      .id = "Workbook") %>%
          dplyr::mutate(filter_buttons = dplyr::case_when(!is.na(filter_buttons) ~ filter_buttons,
                                                          TRUE ~ "false"))
        
      })
      
      
      ##Compile table of font styles
      font_styles <- reactive({
        
        font_style <- function(xml){
          
          text_prop <- xml2::xml_find_all(xml, "//style:text-properties")
          
          tibble(
            
            "font" = text_prop %>% 
              xml2::xml_attr("font-name"),
            
            "font_size" = text_prop %>% 
              xml2::xml_attr("font-size"),
            
            "font_style" = text_prop %>% 
              xml2::xml_attr("font-style"),
            
            "underline" = text_prop %>% 
              xml2::xml_attr("text-underline-type"),
            
            "colour" = text_prop %>% 
              xml2::xml_attr("color"),
            
            "superscript" = text_prop %>%
              xml2::xml_attr("text-position"),
            
            "cell_style" = text_prop %>% 
              xml2::xml_parent() %>%
              xml2::xml_attr("name")
            )
          }
        
        
        purrr::map_df(.x = book_content(),
                      .f = font_style, 
                      .id = "Workbook") %>%
          dplyr::left_join(
            dplyr::select(tidy_format(), 
                        Workbook, sheet, address, cell_style))
      })
      
      ##Compile table of cell styles
      cell_styles <- reactive({
        
        cell_style <- function(xml){
          
          tibble(
            
            "background" = xml2::xml_find_all(xml, "//style:table-cell-properties") %>% 
              xml2::xml_attr("background-color"),
            
            "wrap" = xml2::xml_find_all(xml, "//style:table-cell-properties") %>% 
              xml2::xml_attr("wrap-option"),
            
            "rotation" = xml2::xml_find_all(xml, "//style:table-cell-properties") %>% 
              xml2::xml_attr("rotation-angle"),
            
            "cell_style" = xml2::xml_find_all(xml, "//style:table-cell-properties") %>% 
              xml2::xml_parent() %>%
              xml2::xml_attr("name")
          )
        }
        
        
        purrr::map_df(.x = book_content(),
                      .f = cell_style, 
                      .id = "Workbook") %>%
          ##Join it to the cell names
          dplyr::left_join(dplyr::select(tidy_format(), 
                                         Workbook, sheet, address, cell_style, cell_content))
      })
      
      ##Compile table of stylistic styles
      
      format_styles <- reactive({
        
        table_data <- function(xml){
          table <- xml2::xml_find_all(xml, "//style:style")
          
          tibble(
            "cell_style" = table %>%
              xml2::xml_attr("name") %>%
              null_s(),
            "heading_style" = table %>%
              xml2::xml_attr("parent-style-name") %>%
              null_s(),
            ##Get name of associated data styles
            "data_style" = table %>%
              xml2::xml_attr("data-style-name") %>%
              null_s()
          ) 
        }
        
        purrr::map_df(.x = book_content(),
                      .f = table_data, 
                      .id = "Workbook") 
        
        
      })
      
      ##Create output tables----------------------------------------------------
      
      ##There should be a cover sheet for each table---------------
      output$cover_sheet <- DT::renderDataTable({
        
        workbook_sheet_names() %>%
          #Find candidates for being the cover sheet
          dplyr::mutate(sheet = dplyr::case_when(grepl("cover", sheet, ignore.case = TRUE) ~ sheet,
                                                 TRUE ~ "")) %>%
          dplyr::select(Workbook, sheet) %>%
          unique() %>%
          dplyr::mutate(check = sheet != "") %>%
          ##Keep either blank or cover sheet depending on whether there is a cover sheet
          dplyr::group_by(Workbook) %>%
          dplyr::filter((sheet != "") | (mean(check) == 0 & sheet == "")) %>%
          datatbl()
        
      })
      
      ##There should be a table of contents for each table--------------
      output$contents_sheet <- DT::renderDataTable({
        
        workbook_sheet_names() %>%
          #Find candidates for being the content sheet
          dplyr::mutate(sheet = dplyr::case_when(grepl("content", sheet, ignore.case = TRUE) ~ sheet,
                                                 TRUE ~ "")) %>%
          dplyr::select(Workbook, sheet) %>%
          unique() %>%
          dplyr::mutate(check = sheet != "") %>%
          ##Keep either blank or content sheet depending on whether there is a content sheet
          dplyr::group_by(Workbook) %>%
          dplyr::filter((sheet != "") | (mean(check) == 0 & sheet == "")) %>%
          datatbl()
        
      })
      
      ##There should be a notes sheet for each table--------------
      output$notes_sheet <- DT::renderDataTable({
        
        workbook_sheet_names() %>%
          #Find candidates for being the notes sheet
          dplyr::mutate(sheet = dplyr::case_when(grepl("note", sheet, ignore.case = TRUE) ~ sheet,
                                                 TRUE ~ "")) %>%
          dplyr::select(Workbook, sheet) %>%
          unique() %>%
          dplyr::mutate(check = sheet != "") %>%
          ##Keep either blank or cover sheet depending on whether there is a notes sheet
          dplyr::group_by(Workbook) %>%
          dplyr::filter((sheet != "") | (mean(check) == 0 & sheet == "")) %>%
          datatbl()
        
      })
      
      #The workbook should have a title--------------------
      output$book_title <- DT::renderDataTable({
        
        extract_name <- function(file){
          
          xml2::xml_find_all(file, "//dc:title") %>%
            xml2::as_list() %>%
            unlist() %>%
            null_s()
        }
        
        purrr::map_df(.x = book_meta(),
                      .f = extract_name,
                      .id = "file")  %>%
          tidyr::pivot_longer(cols = dplyr::everything(),
                              names_to = "Workbook",
                              values_to = "Workbook title") %>%
          dplyr::mutate(check = `Workbook title` != "") %>%

          datatbl()
        
        
      })
      
      ###Worksheets should not include images-----
      output$no_images <- DT::renderDataTable({
        
        find_images <- function(xml){
          
          tibble(img =  xml2::xml_find_all(xml, "//draw:image") %>%
                   xml2::xml_attr("href") %>%
                   null_s())
          
        }
        
        purrr::map_df(.x = book_content(),
                      .f = find_images, 
                      .id = "Workbook") %>%
          dplyr::mutate(check = img == "") %>%
          
          datatbl()
        
          
        
      })

      ##All links should have a meaningful name---------------
      output$named_links <- DT::renderDataTable({
        
          find_links <- function(xml){
            
            links <- xml2::xml_find_all(xml, "//text:a")
            
            tibble(
              "link_text" =  links %>%
                xml2::xml_contents() %>%
                xml2::as_list() %>%
                unlist() %>%
                null_s(),
              "link_url" = links %>%
                xml2::xml_attr("href") %>%
                null_s())
          }
          
          
          ##Create table from all sheets
          purrr::map_df(.x = book_content(),
                        .f = find_links, 
                        .id = "Workbook") %>%
            dplyr::mutate(check = link_text != link_url | link_url == "") %>%
            
            datatbl()
      })
      
      ##There should only be one table per sheet------------------
      
    output$single_table <- DT::renderDataTable({
      
      book_tables() %>%
        dplyr::group_by(Workbook, sheet) %>%
        dplyr::summarise(number_of_tables = n()) %>%
        dplyr::mutate(check = number_of_tables == 1) %>%
        datatbl()
      
    })
      
      ##Tables should start in column A-------------------
      
      output$col_a <- DT::renderDataTable({
        
       book_tables() %>%
          #Keep only columns of interest
          dplyr::select(Workbook, sheet, table_name, start_cell) %>%
          dplyr::mutate(check = grepl("^A\\d", start_cell)) %>%
          datatbl()
      })
      
      ##Content should be in marked up tables
      output$no_table <- DT::renderDataTable({
      
      tidy_format() %>%
          dplyr::filter(is_empty == FALSE) %>%
          dplyr::select(Workbook, sheet) %>%
          unique() %>%
          #Join to list of tables
          dplyr::left_join(select(book_tables(),
                                  Workbook, sheet, table_name)) %>%
          dplyr::mutate(check = !is.na(table_name)) %>%
          datatbl()
        
      })
      
      ##Tables should have a meaningful name-------------------
      
      output$named_tables <- DT::renderDataTable({
        
        book_tables() %>%
          #Keep only columns of interest
          dplyr::select(Workbook, sheet, table_name) %>%
          ##Does table name start with something other than table
          dplyr::mutate(check = !grepl("^table", table_name, ignore.case = TRUE)) %>%
          datatbl()
      })
      
      #Tables should have filter buttons switched off---------------------
      
      output$hidden_filter <- renderDataTable({
        
        book_tables() %>%
        #Keep only columns of interest
          dplyr::select(Workbook, sheet, table_name, filter_buttons) %>%
          dplyr::mutate(check = filter_buttons != "true") %>%
            datatbl()
        
      })
      
      ##Font should be Arial-------------
      output$font_family <- DT::renderDataTable({
        
        font_styles() %>%
          dplyr::group_by(Workbook, sheet, font) %>%
          dplyr::summarise(cells = paste(address, collapse = ", "),
                           count_of_cells = n()) %>%
          ##Remove anything that's Arial, we don't need to list every cell
          dplyr::filter(!is.na(sheet)) %>%
          dplyr::mutate(cells = case_when(font == "Arial" ~ "multiple cells",
                                          TRUE ~ cells),
                        check = cells == "multiple cells") %>%
          datatbl()
        
      })
      
      ##All text should be wrapped and visible
      output$wrapped <- DT::renderDataTable({
        
        cell_styles() %>%
          #Drop all the empty cells
          dplyr::filter(!is.na(cell_content)) %>%
          #Calculate how long content is
          dplyr::mutate(length = nchar(cell_content),
                        check = wrap == "wrap") %>%
          dplyr::filter(length > 100) %>%
          dplyr::select(Workbook, sheet, address, length, check) %>%
          datatbl()
        
      })
      
      output$headings <- DT::renderDataTable({
        
        ##Get tidy data just for cell A1
        tidy_format() %>%
          dplyr::filter(address == "A1") %>%
          dplyr::select(Workbook, sheet, address, cell_style) %>%
          ##Join to styles
          dplyr::left_join(
            format_styles()) %>%
          dplyr::select(-cell_style, -data_style) %>%
          dplyr::mutate(check = grepl("heading", heading_style, ignore.case = TRUE)) %>%
          datatbl()
      })
      
      ##Font should be minimum size 10-----------------
      output$font_size <- DT::renderDataTable({
        
        font_styles() %>%
          dplyr::group_by(Workbook, sheet, font_size) %>%
          dplyr::summarise(cells = paste(address, collapse = ", "),
                           count_of_cells = n()) %>%
          ##Extract numeric values and set default
          dplyr::mutate(
            numeric_font = as.numeric(gsub("^(\\d{1,3})pt", "\\1", font_size)),
            font_size = dplyr::case_when(is.na(numeric_font) ~ "default",
                                         TRUE ~ font_size)) %>%
          dplyr::filter(!is.na(sheet)) %>%
          dplyr::mutate(cells = case_when(is.na(numeric_font) | numeric_font >= 10 ~ "multiple cells",
                                          TRUE ~ cells),
                        check = is.na(numeric_font) | numeric_font >= 10) %>%
          dplyr::select(-numeric_font) %>%
          datatbl()
          
      })
      
      ##Font colour should be automatic
      output$font_colour <- DT::renderDataTable({
        
        font_styles() %>%
          dplyr::group_by(Workbook, sheet, colour) %>%
          dplyr::summarise(cells = paste(address, collapse = ", "),
                           count_of_cells = n()) %>%
          dplyr::filter(!is.na(sheet)) %>%
          dplyr::mutate(cells = case_when(is.na(colour) ~ "multiple cells",
                                          TRUE ~ cells),
                        check = is.na(colour),
                        colour = case_when(is.na(colour) ~ "automatic",
                                           TRUE ~ colour)) %>%
          datatbl()
        
      })
      
      
      ##Cells should have no background fill-----------------
      output$background_fill <- DT::renderDataTable({
        
        cell_styles() %>%
          dplyr::group_by(Workbook, sheet, background) %>%
          dplyr::summarise(cells = paste(address, collapse = ", "),
                           count_of_cells = n()) %>%
          dplyr::filter(!is.na(sheet)) %>%
          dplyr::mutate(cells = case_when(is.na(background) ~ "multiple cells",
                                          TRUE ~ cells),
                        check = is.na(background)) %>%
          datatbl()
        
      })
      
      ##All text should be horizontal---------
      output$text_horizontal <- DT::renderDataTable({
        
        cell_styles() %>%
          dplyr::group_by(Workbook, sheet, rotation) %>%
          dplyr::summarise(cells = paste(address, collapse = ", "),
                           count_of_cells = n()) %>%
          dplyr::filter(!is.na(sheet)) %>%
          dplyr::mutate(cells = case_when(is.na(rotation) ~ "multiple cells",
                                          TRUE ~ cells),
                        check = is.na(rotation),
                        rotation = dplyr::case_when(is.na(rotation) ~ "horizontal",
                                                    TRUE ~ rotation)) %>%
          datatbl()
        
      })
      
      ##There should be no italics/underlined text
      output$italic_underline <- DT::renderDataTable({
        font_styles() %>%
          dplyr::group_by(Workbook, sheet, font_style, underline) %>%
          dplyr::summarise(cells = paste(address, collapse = ", "),
                           count_of_cells = n()) %>% 
          dplyr::filter(!is.na(sheet)) %>%
          dplyr::mutate(cells = case_when(is.na(font_style) & is.na(underline) ~ "multiple cells",
                                          TRUE ~ cells),
                        check = is.na(font_style) & is.na(underline)) %>%
          datatbl()
        
      })
    
      
      #There should be no empty worksheets---------------------------
      
      output$empty_sheets <- DT::renderDataTable({
        
        tidy_format() %>%
          dplyr::group_by(Workbook, "Sheet name" = sheet) %>%
          dplyr::summarise("Filled cells" = sum(is_empty == FALSE)) %>%
          dplyr::mutate(check = `Filled cells` != 0) %>%
          datatbl()
        
        
      })
      
      #There should be no hidden columns---------------------------
      
      output$hidden_cols <- DT::renderDataTable({
        
       ##Function to check for hidden stuff
        
        hidden_stuff <- function(xml){
          
          
          tibble("sheet_name" = 
                   xml2::xml_find_all(xml, "//table:table") %>%
                   xml2::xml_attr("name")) %>%
            
            dplyr::left_join(
              
              tibble("sheet_name" = 
                       xml2::xml_find_all(xml, "//table:table-column[@table:visibility='collapse']") %>%
                       xml2::xml_parent() %>%
                       xml2::xml_attr("name")) %>%
                dplyr::mutate("hidden_column" = TRUE)
            ) %>%
            
            dplyr::left_join(
              tibble("sheet_name" = 
                       xml2::xml_find_all(xml, "//table:table-row[@table:visibility='collapse']") %>%
                       xml2::xml_parent() %>%
                       xml2::xml_attr("name")) %>%
                dplyr::mutate("hidden_row" = TRUE)
            )
          
        }
        
        #Map it over every workbook
        purrr::map_df(.x = book_content(),
                      .f = hidden_stuff,
                      .id = "Workbook") %>%
          #drop weird sheet names
          dplyr::filter(!grepl("file://", sheet_name, fixed = TRUE)) %>%
        
          dplyr::mutate(check = is.na(hidden_column) & is.na(hidden_row)) %>%
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
      
      
   
      #No empty rows and columns--------------------
      
      output$empty_cols <- DT::renderDataTable({
        
        ##Summarise empty rows
        empty_rows <- tidy_format() %>%
          ##Group by cols and count full cells
          dplyr::group_by(Workbook, sheet, row) %>%
          dplyr::summarise(filled_rows = sum(is_empty == FALSE))
        
        ##Remove rows beyond last filled row
        max_row <- empty_rows %>%
          dplyr::filter(filled_rows != 0) %>%
          dplyr::group_by(Workbook, sheet) %>%
          dplyr::summarise(end_row = max(row, na.rm = TRUE))

        empty_rows <- empty_rows %>%
          dplyr::left_join(max_row, by = c("Workbook", "sheet")) %>%
          dplyr::filter(row <= end_row) %>%
          dplyr::select(-end_row) %>%
          ##Join filled rows data
          ##Produce a list of empty cols
          dplyr::group_by(Workbook, sheet) %>%
          dplyr::filter(filled_rows == 0) %>%
          dplyr::summarise(empty_rows = paste(row, collapse = ","))
        
        ##Summarise empty cols
        empty_cols <- tidy_format() %>%
          ##Group by cols and count full cells
          dplyr::group_by(Workbook, sheet, col) %>%
          dplyr::summarise(filled_cols = sum(is_empty == FALSE)) 
          
        ##Remove cols beyond last filled column
        max_col <- empty_cols %>%
          dplyr::filter(filled_cols != 0) %>%
          dplyr::group_by(Workbook, sheet) %>%
          dplyr::summarise(end_col = max(col, na.rm = TRUE))
        
        empty_cols %>%
          dplyr::left_join(max_col, by = c("Workbook", "sheet")) %>%
          dplyr::filter(col <= end_col) %>%
          dplyr::select(-end_col) %>%

          ##Join filled rows data
          ##Produce a list of empty cols
          dplyr::group_by(Workbook, sheet) %>%
          dplyr::filter(filled_cols == 0) %>%
          dplyr::summarise(empty_cols = paste(col, collapse = ",")) %>%
          dplyr::full_join(empty_rows, by = c("Workbook", "sheet")) %>%

          dplyr::mutate(check = FALSE) %>%
          datatbl()

      })
      
      #No merged cells--------------------
      
      output$merged_cells <- DT::renderDataTable({
        
        tidy_format() %>%
          ##Count how many are merged
          dplyr::group_by(Workbook, sheet) %>%
          dplyr::summarise(number_of_merged_cells = sum(is_merged)) %>%
          dplyr::mutate(check = number_of_merged_cells == 0) %>%
          datatbl()
        
      })
      
     
      ##Cells that contain superscript------------------------------------
      
      output$superscript <- DT::renderDataTable({
        
       
        font_styles() %>%
          #Remove cells with no superscript
          dplyr::filter(!is.na(superscript)) %>%
          dplyr::filter(!is.na(sheet)) %>%
          dplyr::mutate(superscript = dplyr::case_when(superscript == "-33%" ~ "subscript",
                                                       superscript == "33%" ~ "superscript")) %>%
          dplyr::group_by(Workbook, sheet, superscript) %>%
          dplyr::summarise(cells = paste(address, collapse = ", "),
                           count_of_cells = n(),
                           check = FALSE) %>%
          datatbl()
        
      })
      
      ##Date style should be day month year--------------
      output$date_style <- DT::renderDataTable({
        
        ##Loop function over all xml files
        all_xml <- function(xml){
          
          ##Loop over every different date style
          date_order <- function(date_style) {
            
            out <- date_style %>% 
              xml_children() %>%
              xml2::xml_name() 
            
            out <- out[out %in% c("day", "month", "year")]
            
            tibble(
              "data_style" = date_style %>% 
                xml2::xml_attr("name"),
              
              ##Drop anything that isn't day-month-year and collapse into a string
              "date_order" = paste(out, collapse = "-")
            )
            
          }
          
          purrr::map_df(.x = xml2::xml_find_all(xml, 
                                                "//number:date-style"),
                        .f = date_order) 
        }
        
        ##Map over all files provided
        data <- purrr::map_df(.x = book_style(),
                      .f = all_xml,
                      .id = "Workbook")
          
          
        if(nrow(data) == 0){
          tibble(Workbook = character(), 
                 sheet = character(), 
                 address = character(), 
                 cell_style = character())
          
        } else{
        
        data %>%
          ##Join to get names of the cell styles, not just the number styles
          dplyr::left_join(format_styles()) %>%
          dplyr::select(-heading_style) %>%
          ##Join to tidy format to get cell addresses
          dplyr::left_join(
            dplyr::select(tidy_format(),
                          Workbook, sheet, address, cell_style)) %>%
          ##Group up values
          dplyr::group_by(Workbook, sheet, date_order) %>%
          dplyr::summarise(cells = paste(address, collapse = ", "),
                           count_of_cells = n()) %>%
          dplyr::mutate(cells = 
                          dplyr::case_when(date_order == "day-month-year" ~ "multiple cells",
                                           TRUE ~ cells),
                        check = date_order == "day-month-year") %>%
          datatbl()
        }
        
      })
      
    
      output$comma_format <- DT::renderDataTable({
        
        tidy_format() %>%
          #Only keep rows this is relevant to
          dplyr::filter(value_type == "float" & numeric_value >= 1000) %>%
          dplyr::select(Workbook, sheet, address, cell_content, base_value) %>%
          dplyr::mutate(check = grepl("[,]\\d{3}", cell_content)) %>%
          dplyr::filter(check == FALSE) %>%
          datatbl()
      })
      
      ##Help message popup-----------------------------
      observeEvent(input$show, {
        showModal(modalDialog(
          title = "How to use the accessibility checker",
          "Click the 'Browse' button to upload one or more publication tables
      in .ods format into the file input. 
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