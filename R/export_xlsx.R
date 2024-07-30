#' Add response to single survey question
#'
#' @param data survey data, formatted where each row is an individual and column names describe demographics or survey questions (of the format YY__ZZ, where XX is the theme and YY the question, where ' ' are replaced with '_' and '?' with 'XX'.)
#' @param question_column the column index of the question within data
#' @param demographic_columns column indices of demographics used to filter the data (e.g. sex, etc)
#' @param append_to Flag (TRUE/FALSE) for whether the summary sheet is appended to a previous workbook
#' @param wb the workbook to append the sheet to (requires append_to = TRUE)
#' @param sheet the sheet name
#'
#' @return
#' @export
to_sheet_single_survey_question <- function(data,
                                            question_column,
                                            demographic_columns,
                                            append_to = NULL,
                                            wb = NULL,
                                            sheet = 'Data'){
  ##########
  ### 1) Get the question from the column name.
  ##########
  theme_question = names(data)[question_column] |> stringr::str_split('__') |> unlist()
  theme = theme_question[1] ; question = theme_question[2]
  theme = theme |> stringr::str_replace_all('_',' ') |> stringr::str_replace_all('XX','?')
  question = question |> stringr::str_replace_all('_',' ') |> stringr::str_replace_all('XX','?')

  ##########
  ### 2) Format the data into columns = deomgraphics, rows = total count + percent answers.
  ##########
  question_column_text = names(data)[question_column]
  question_values = data[[question_column_text]] |> levels()
  demographic_columns_text = names(data)[demographic_columns]

  all = list(list(demo = 'all', value = NA))
  demo_values = lapply(demographic_columns_text, function(demo_col){
    if(is.factor(data[[demo_col]])){
      data[[demo_col]] = data[[demo_col]] |> as.character()
    }
    values = data[[demo_col]] |> unique()
    values = values[!is.na(values)]
    out = list()
    for(i in 1:length(values))
      out[[i]] = list(demo = demo_col, value = values[i])
    return(out)
  })
  demo_values = c(all, do.call(c, demo_values))
  demo_values_names = lapply(demo_values, function(x){paste0(x, collapse = ': ')}) |> unlist()
  demo_values_names[1] = 'Total'
  data_formatted = lapply(demo_values, function(x){
    demo = x$demo ; value = x$value
    if(demo == 'all'){
      count = data[[question_column_text]] |>table()
    }else{
      count =  data[[question_column_text]][which(data[[demo]] == value)] |>table()
    }

    total = sum(count)
    values = rep(0, length(question_values))
    values[match(names(count), question_values)] = count |> as.numeric()
    percent = values/total
    return(c(total, percent))
  }) |> data.frame() #|> t() |> data.frame()
  row.names(data_formatted) = c('Total', question_values)
  colnames(data_formatted) = demo_values_names

  #Sneaky add empty columns around demo groups.
  demo = lapply(demo_values, function(x){x[[1]]}) |> unlist()
  d = data_formatted
  cur_name = demo[ncol(data_formatted)]
  for(i in ncol(data_formatted):2){
    if(cur_name != demo[i-1]){
      names_og = names(data_formatted)
      names_a  =names_og[1:(i-1)] ; names_b = names_og[(i):ncol(data_formatted)]
      data_formatted = data.frame(data_formatted[1:(i-1)],
                                  '',
                                  data_formatted[(i):ncol(data_formatted)]
      )
      names(data_formatted) = c(names_a, '', names_b)
    }
    cur_name =  demo[i-1]
  }

  demo = lapply(names(data_formatted), function(x){(x |> stringr::str_split(': ') |> unlist())[1]}) |>
    unlist()  |>
    stringr::str_replace_all('__',': ') |>
    stringr::str_replace_all('_',' ') |>
    stringr::str_replace_all('XX','?')
  value = lapply(names(data_formatted), function(x){(x |> stringr::str_split(': ') |> unlist())[2]}) |> unlist()
  data_formatted_header =  data.frame(demo =demo,
                                      value = value ) |> t() |> data.frame()
  names(data_formatted_header) = paste0('h',1:ncol(data_formatted_header))


  ##########
  ### 3) Create and format the excel sheet
  ##########
  # Create and save the excel file
  if(!is.null(append_to)){
    wb <-  append_to
  }else if(is.null(wb)){
    wb <- openxlsx::createWorkbook()
  }
  #styles
  top_border = openxlsx::createStyle(border = 'top', borderColour = 'black', borderStyle = 'thin')
  top_bottom_border = openxlsx::createStyle(border = 'TopBottom', borderColour = 'black', borderStyle = 'thin')
  bottom_border = openxlsx::createStyle(border = 'bottom', borderColour = 'black', borderStyle = 'thin')


  openxlsx::addWorksheet(wb, sheet, gridLines = FALSE)
  start_row = 2

  #Add OME logo to top left
  openxlsx::insertImage(wb, sheet = sheet,file = paste0(system.file(package = "OMEStools"), "/logo.png"), width = 700/400, height = 228/400, startRow = 2, startCol = 1)
  # Add question + theme(title)
  openxlsx::writeData(wb, sheet = sheet, x = 'Theme:', startRow = start_row, startCol = 3)
  openxlsx::writeData(wb, sheet = sheet, x = theme, startRow = start_row, startCol = 4)
  openxlsx::writeData(wb, sheet = sheet, x = 'Question:', startRow = start_row+1, startCol = 3)
  openxlsx::writeData(wb, sheet = sheet, x = question, startRow = start_row+1, startCol = 4)
  main_title = openxlsx::createStyle(fontSize = 12, textDecoration = "bold", halign = "left")
  openxlsx::addStyle(wb = wb, sheet = sheet, rows = 2:3, cols = 3, style = main_title)
  # Add data (header)
  openxlsx::writeData(wb,
                      sheet = sheet,
                      x = data_formatted_header,
                      startRow = start_row+3,
                      startCol = 3,
                      rowNames = FALSE,
                      colNames = FALSE,
                      withFilter = FALSE)
  # Merge the demographics with repeated values
  demo_cat = data_formatted_header[1,] |> as.character()
  demo_cat_unique = demo_cat  |> unique()
  demo_kind_style =
    for(i in 1:length(demo_cat_unique)){
      if(demo_cat_unique[i] == ''){
        next
      }
      match_index = which(demo_cat  == demo_cat_unique[i])
      pos = min(match_index):max(match_index)
      if(length(pos) > 1){
        openxlsx::mergeCells(wb,sheet = sheet,cols = pos+2, rows = start_row+3)
        openxlsx::addStyle(wb,sheet = sheet,cols = pos+2, rows = start_row+3, style = bottom_border,stack = T)
      }
    }
  table_head_style = openxlsx::createStyle(halign = 'center', valign = 'center', wrapText = T)
  openxlsx::addStyle(wb = wb, sheet = sheet,
                     rows = 0:1 + (start_row+3),
                     cols = 3:(3+ncol(data_formatted_header)),
                     style = table_head_style,
                     stack = T,
                     gridExpand = TRUE)


  # Add data (body)
  openxlsx::writeData(wb,
                      sheet = sheet,
                      x = data_formatted,
                      startRow = start_row + 5,
                      startCol = 2,
                      rowNames = T,
                      colNames = FALSE,
                      withFilter = FALSE)

  # Format the percentage values (with 0dp)
  style = openxlsx::createStyle(numFmt = "0%")
  openxlsx::addStyle(wb = wb, sheet = sheet,
                     rows = 1:(nrow(data_formatted)-1) + start_row + 5,
                     cols = 1:(ncol(data_formatted)+1)+1,
                     style = style,
                     gridExpand = TRUE)
  # Set row heights + colun widths
  openxlsx::setRowHeights(wb, sheet = sheet,
                          rows = c(6),
                          heights = c(40))
  openxlsx::setColWidths(wb, sheet = sheet,
                         cols = c(2),
                         widths = c(24.8))
  which(demo_cat == '')+2
  openxlsx::setColWidths(wb, sheet = sheet,
                         cols = c(which(demo_cat == '')+2),
                         widths = c(1.33))
  # Freeze the first two columns for scrolling.
  openxlsx::freezePane(wb,sheet = sheet,firstActiveCol = 3)
  # Add borders
  openxlsx::addStyle(wb = wb, sheet = sheet,
                     rows = start_row +3,
                     cols = 2:(2+ncol(data_formatted)),
                     style = top_border, stack = T)
  openxlsx::addStyle(wb = wb, sheet = sheet,
                     rows = start_row +4 + nrow(data_formatted),
                     cols = 2:(2+ncol(data_formatted)),
                     style = bottom_border, stack = T)
  openxlsx::addStyle(wb = wb, sheet = sheet,
                     rows = start_row +5,
                     cols = 2:(2+ncol(data_formatted)),
                     style = top_bottom_border, stack = T)
  openxlsx::addStyle(wb = wb, sheet = sheet,
                     rows = 4:(start_row +4 +nrow(data_formatted)),
                     cols = 1:(2+ncol(data_formatted)),
                     style = table_head_style, stack = T,gridExpand = TRUE)
  openxlsx::pageBreak(wb, sheet = sheet, i = 4, type = "column")
  # openxlsx::saveWorkbook(wb, file = 'working.xlsx', overwrite = T)

  return(wb)
}

#' Add theme summary sheet
#'
#' @inheritParams to_sheet_single_survey_question
#' @param theme string containing the theme name.
#' @param theme_columns Indices of the columns in the data for the theme
#' @param order not used
#'
#' @return a workbook with the theme summary added
#' @export
to_sheet_theme_summary <- function(data,
                                   theme = NULL,
                                   theme_columns = NULL,
                                   append_to = NULL,
                                   wb = NULL,
                                   sheet = 'Data',
                                   order = FALSE){

  # 1) Get the column indices of the data relating to the theme.
  if(!is.null(theme)){
    theme = theme |>stringr::str_replace_all(' ','_')
    theme_columns = grep(theme, names(data))
  }
  if(is.null(theme_columns)){
    stop('Theme issue.')
  }

  df = data[,theme_columns]
  theme_question = names(df) |> stringr::str_split('__')
  theme = lapply(theme_question, function(x){x[1]})[[1]] |> stringr::str_replace_all('_',' ') |> stringr::str_replace_all('XX','?')
  questions = lapply(theme_question, function(x){x[2]}) |> stringr::str_replace_all('_',' ') |> stringr::str_replace_all('XX','?')
  question_values = df[,1] |> levels()


  counter = 1
  data_formatted = lapply(df, function(x){
    if(is.factor(x)){
      x = x |> as.character()
    }
    count = x[!is.na(x)]|>table()

    total = sum(count)
    values = rep(0, length(question_values))
    values[match(names(count), question_values)] = count |> as.numeric()
    percent = values/total
    return(c(total, percent))
  }) |> data.frame() #|> t() |> data.frame()
  row.names(data_formatted) = c('Total', question_values)
  colnames(data_formatted) = questions


  ##########
  ### 3) Create and format the excel sheet
  ##########
  # Create and save the excel file
  if(!is.null(append_to)){
    wb <-  append_to
  }else if(is.null(wb)){
    wb <- openxlsx::createWorkbook()
  }
  #styles
  top_border = openxlsx::createStyle(border = 'top', borderColour = 'black', borderStyle = 'thin')
  top_bottom_border = openxlsx::createStyle(border = 'TopBottom', borderColour = 'black', borderStyle = 'thin')
  bottom_border = openxlsx::createStyle(border = 'bottom', borderColour = 'black', borderStyle = 'thin')
  table_head_style = openxlsx::createStyle(halign = 'center', valign = 'center', wrapText = T)

  openxlsx::addWorksheet(wb, sheet, gridLines = FALSE)
  start_row = 2

  #Add OME logo to top left
  openxlsx::insertImage(wb, sheet = sheet,file = paste0(system.file(package = "OMEStools"), "/logo.png"), width = 700/400, height = 228/400, startRow = 2, startCol = 1)
  # Add question + theme(title)
  openxlsx::writeData(wb, sheet = sheet, x = 'Theme:', startRow = start_row, startCol = 3)
  openxlsx::writeData(wb, sheet = sheet, x = theme, startRow = start_row, startCol = 4)
  main_title = openxlsx::createStyle(fontSize = 12, textDecoration = "bold", halign = "left")
  openxlsx::addStyle(wb = wb, sheet = sheet, rows = 2, cols = 3, style = main_title)

  # Add data (body)
  openxlsx::writeData(wb,
                      sheet = sheet,
                      x = data_formatted,
                      startRow = start_row + 3,
                      startCol = 2,
                      rowNames = T,
                      colNames = T,
                      withFilter = FALSE)
  # Wrap column names of data body.
  openxlsx::addStyle(wb = wb, sheet = sheet, rows = 5, cols = 1:(ncol(data_formatted)+1)+1, style = table_head_style)
  # Format the percentage values (with 0dp)
  style = openxlsx::createStyle(numFmt = "0%")
  openxlsx::addStyle(wb = wb, sheet = sheet,
                     rows = 1:(nrow(data_formatted)-1) + start_row + 4,
                     cols = 1:(ncol(data_formatted)+1)+1,
                     style = style,
                     gridExpand = TRUE)
  # Set row heights + colun widths
  openxlsx::setRowHeights(wb, sheet = sheet,
                          rows = c(5),
                          heights = c(50))
  openxlsx::setColWidths(wb, sheet = sheet,
                         cols = c(2,1:ncol(data_formatted)+2),
                         widths = c(24.8, rep(19.8, ncol(data_formatted))))
  # Freeze the first two columns for scrolling.
  openxlsx::freezePane(wb,sheet = sheet,firstActiveCol = 3)
  # Add borders
  openxlsx::addStyle(wb = wb, sheet = sheet,
                     rows = start_row +3,
                     cols = 2:(2+ncol(data_formatted)),
                     style = top_border, stack = T)
  openxlsx::addStyle(wb = wb, sheet = sheet,
                     rows = start_row +3 + nrow(data_formatted),
                     cols = 2:(2+ncol(data_formatted)),
                     style = bottom_border, stack = T)
  openxlsx::addStyle(wb = wb, sheet = sheet,
                     rows = start_row +4,
                     cols = 2:(2+ncol(data_formatted)),
                     style = top_bottom_border, stack = T)
  openxlsx::addStyle(wb = wb, sheet = sheet,
                     rows = 4:(start_row +4 +nrow(data_formatted)),
                     cols = 1:(2+ncol(data_formatted)),
                     style = table_head_style, stack = T,gridExpand = TRUE)
  # openxlsx::saveWorkbook(wb, file = 'working.xlsx', overwrite = T)

  return(wb)
}



#' Add table of contents to an workbook
#'
#' @inheritParams to_sheet_single_survey_question
#' @param link_text character vector of the text displayed for the linked text in the TOC sheet. Must have the same length as the number of sheets in wb.
#' @param return_pos integer pair, of the row and column index to place the return to contents link on the sheets in the wb.
#'
#' @return
#' @export
#'
add_TOC_sheet <- function(wb,
                          link_text = NULL,
                          return_pos = c(1,1)){
  ## A) sanity checks
  sheets = names(wb)
  if(length(sheets) != length(link_text)){
    stop('Stop! The number of sheets does not match the length of the link text provided')
  }

  ## 1) Create TOC sheet.
  openxlsx::addWorksheet(wb, 'Contents', gridLines = F)
  #Add OME logo to top left
  openxlsx::insertImage(wb, sheet = 'Contents',file = paste0(system.file(package = "OMEStools"), "/logo.png"), width = 700/400, height = 228/400, startRow = 2, startCol = 1)
  # Header (Table of Contents)
  openxlsx::writeData(wb, sheet = 'Contents', x = 'Table of Contents', startRow = 2, startCol = 4)
  main_title = openxlsx::createStyle(fontSize = 18, textDecoration = "bold", halign = "center")
  openxlsx::addStyle(wb = wb, sheet = 'Contents', rows = 2, cols = 4, style = main_title)
  # Column widths
  openxlsx::setColWidths(wb, sheet = 'Contents',
                         cols = c(4),
                         widths = c(100 ))
  # Add link text table
  df = data.frame(id = 1:length(link_text), question = link_text)
  names(df) = c('','Individual Tables')
  openxlsx::writeData(wb,
                      sheet = 'Contents',
                      x = df,
                      startRow = 5,
                      startCol = 3,
                      rowNames = F,
                      colNames = T,
                      withFilter = FALSE)
  for(i in 1:length(link_text)){ # Add hyperlinks to other sheets
    openxlsx::writeFormula(wb, "Contents",
                           startRow = 5+i,
                           startCol = 4,
                           x = openxlsx::makeHyperlinkString(sheet = sheets[i], row = 1, col = 1,
                                                             text = link_text[i])
    )
  }
  # Align in the table number to the right.
  right_text = openxlsx::createStyle( halign = "right")
  openxlsx::addStyle(wb = wb, sheet = 'Contents', rows = 1:(length(link_text)+5), cols = 3, style = right_text)


  # Add hyperlink to all other sheets back to the contents sheet.
  for(i in 1:length(sheets)){

    openxlsx::writeFormula(wb, sheets[i],
                           startRow = return_pos[1],
                           startCol = return_pos[2],
                           x = openxlsx::makeHyperlinkString(sheet = 'Contents', row = 1, col = 1,
                                                             text = 'Back to Contents')
    )
  }

  openxlsx::worksheetOrder(wb) = c(length(sheets)+1, 1:length(sheets))
  openxlsx::activeSheet(wb) = 'Contents'
  return(wb)
}
