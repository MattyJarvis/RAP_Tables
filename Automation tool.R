###PLEASE NOTE: This tool is in the early prototype stage. Feedback is welcomed ###
#THe prime aim of this tool is to prive a basic tool to construct tables in R


#dummy data
data <- c( c(1000, 100023432400, 342290328049, 2902))
data <- rbind( data, data, data)
View(data)
######CREATING A HEADER#######

#header function, as the header is standard, most of the process is encapsulated in one function
#we just need to get the table height for the data you are using

table_height <- nrow(data)

#what is the title of the sheet?

title <- c("INSERT TITLE HERE")


#what do you want your sheet to be called?

sheet_name <- c("INSERT SHEET NAME HERE")

#use the 'create_header' function to create a header, it obeys standard convention for Sanctions Tables
#up to ten lines allowed in code, though you can manually edit the code if you want more
#write the code as: data <- create_header(data, date_a, date_b, ...., title vector)
#... = vectors you have created to further amend header, limited to 3 due to standard formatting
#title vector is the character names representing the column names of the table
create_header <- function(data, date_a= c("MISSING DATE!"), date_b = c("MISSING DATE!"), line7 = c(rep("BLANK",ncol(data))), line8= c(rep("BLANK",ncol(data))), line9=  c(rep("BLANK",ncol(data))), tableheader = c("NO TABLE HEADER!", rep("BLANK", sum(ncol(data)-1)))) {  
data <- huxtable::insert_row(data,c(rep("", ncol(data))))
data <- huxtable::insert_row(data, c('Back to Contents', rep(NA, sum(ncol(data)-1))), after = 1)
data <- huxtable::insert_row(data, c(rep("", ncol(data))), after = 2)
data <- huxtable::insert_row(data, c(title, rep(NA, sum(ncol(data)-1))), after = 3)
data <- huxtable::insert_row(data,c(paste0(date_a, " to ", date_b), rep(NA,sum(ncol(data)-1))), after = 4)
data <- huxtable::insert_row(data, c(rep("", ncol(data))), after = 5)
data <- huxtable::insert_row(data, line7, after =6)
data <- huxtable::insert_row(data, line8, after =7)
data <- huxtable::insert_row(data, line9, after =8)
data <- huxtable::insert_row(data, tableheader, after = 9)
data <- subset(data, data[,1] != "BLANK")
data <- subset(data, data[,1] != TRUE)
}

data <- create_header(data)


#need to take height of header for code later

header_height<- nrow(data) - table_height


#######CREATING A FOOTER######

#use the 'create_footer' command to add a footer to your table
#up to ten lines allowed in code, though you can manually edit the code if you want more
#write the code as: data <- create_header(data, ...)
#... = vectors you have created to further amend footer
create_footer <- function(data, line1 = c(rep("BLANK",ncol(data))), line2 = c(rep("BLANK",ncol(data))), line3 = c(rep("BLANK",ncol(data))), line4 = c(rep("BLANK",ncol(data))), 
                                                                            line5 = c(rep("BLANK",ncol(data))), line6 = c(rep("BLANK",ncol(data))), 
                                                                            line7 = c(rep("BLANK",ncol(data))), line8 = c(rep("BLANK",ncol(data))), 
                                                                            line9 = c(rep("BLANK",ncol(data))), line10 = c(rep("BLANK",ncol(data)))){
  data <- huxtable::insert_row(data, line1, after = nrow(data))
  data <- huxtable::insert_row(data, line2, after = nrow(data))
  data <- huxtable::insert_row(data, line3, after = nrow(data))
  data <- huxtable::insert_row(data, line4, after = nrow(data))
  data <- huxtable::insert_row(data, line5, after = nrow(data))
  data <- huxtable::insert_row(data, line6, after = nrow(data))
  data <- huxtable::insert_row(data, line7, after = nrow(data))
  data <- huxtable::insert_row(data, line8, after = nrow(data))
  data <- huxtable::insert_row(data, line9, after = nrow(data))
  data <- huxtable::insert_row(data, line10, after = nrow(data))
  data <- subset(data, data[,1] != "BLANK")
  data <- subset(data, data[,1] != "TRUE")
  }


data <- create_footer(data, c(rep("HELLo!", ncol(data))), c(rep("Byebye", ncol(data))))
View(data)
######CREATING AND FORMATTING A SHEET#####

#we first need to create the document
#use `wb <- createWorkbook()` if this is the first part of the excel
#you then need to add a sheet using 'addWorksheet(wb, sheetname = sheet_title)
#if you have already created the workbook, you just need to follow the add sheet step

wb <- createWorkbook()

addWorksheet(wb, sheetName = sheet_name)

#set current sheet as the sheet number we are working on
currentsheet <- 1

#styles are created to format cells in the way we want, this also sets borders to be invivisble
#this is using the 'openxlsx::createStyle' function, use this is you want to create more custom formats
noborder<- openxlsx::createStyle(fontName = NULL, fontSize=NULL, fontColour = NULL, wrapText = FALSE, border = "topbottomleftright" , borderColour = "white")
numwithcomma <- openxlsx::createStyle(fontName = NULL, fontSize=NULL, fontColour = "NULL", numFmt= "COMMA", wrapText = FALSE)
titleformat <- openxlsx::createStyle(fontName = NULL, fontSize=16, fontColour = NULL, wrapText = FALSE, textDecoration = 'bold')
subtitleformat <- openxlsx::createStyle(fontName = NULL, fontSize = 14, fontColour = NULL, wrapText = FALSE, textDecoration = 'bold')
general_bold <- createStyle(fontName = NULL, fontColour = NULL, wrapText = FALSE, textDecoration = 'bold')
tableheaderformat <- createStyle(fontName = NULL, fontColour =  NULL, wrapText = TRUE, textDecoration = 'bold', border = "bottom")
tablecolumnclassifier <- createStyle(fontName = NULL, fontColour = NULL, wrapText = FALSE, textDecoration = 'bold', border = "bottom")

View(data)

#inputting and formating the table
#to write the data, you just need to input the data and sheet (name or number) you want
#the import table function will take into account header height to place your data in the correct place
#it will also apply formatting, which can be further customised (see reference document)
import_table <- function (data, sheet = currentsheet, rowcomma = sum(header_height+1):sum(header_height+table_height), colcomma = 1:ncol(data), rowbold = sum(header_height+1):sum(header_height+table_height), colbold= 1:ncol(data)){
tabledata <- data[sum(header_height+2):sum(header_height+table_height), 1:ncol(data)]

tabledata[,1] <- as.numeric(as.character(tabledata[,1]))
tabledata[,2] <- as.numeric(as.character(tabledata[,2]))
tabledata[,3] <- as.numeric(as.character(tabledata[,3]))
tabledata[,4] <- as.numeric(as.character(tabledata[,4]))
writeData(wb, tabledata, sheet = sheet, startRow = sum(header_height+1), colNames = FALSE)
openxlsx::addStyle(wb, sheet = sheet, style = noborder, rows = 1:sum(nrow(data)+100), cols = 1:sum(ncol(data)+100), gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet = sheet, style = numwithcomma, rows = rowcomma, cols = sum(colcomma), gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet = sheet, style = general_bold, rows = rowbold, cols = colbold, gridExpand = TRUE, stack = TRUE)
}
tabledata <- data[sum(header_height+2):sum(header_height+table_height), 1:ncol(data)]


tabledata[,1] <- as.numeric(as.character(tabledata[,1]))
tabledata[,2] <- as.numeric(as.character(tabledata[,2]))
tabledata[,3] <- as.numeric(as.character(tabledata[,3]))
import_table(data, sheet = 1, rowbold = FALSE, colbold = FALSE)
summary(tabledata)

#inputting and formating the header
#to write the data, you just need to input the data and sheet you want
#it will also apply formatting, which that can be customised (see reference document)
import_header <- function (data, sheet= currentsheet, rowbold = 1:header_height, colbold= 1:ncol(data), rowtitle = 4, coltitle = 1, rowsubt = 5, colsubt = 1, rowtablehead = header_height, coltablehead = 1:ncol(data)){
  data <- data[1:header_height,]
  writeData(wb, sheet = sheet, data, startRow = 1, colNames = FALSE)
  openxlsx::addStyle(wb, sheet = sheet, style = titleformat, rows = rowtitle, cols = coltitle, gridExpand = TRUE, stack = TRUE)
  openxlsx::addStyle(wb, sheet = sheet, style = subtitleformat, rows = rowsubt, cols = colsubt, gridExpand = TRUE, stack = TRUE)
  openxlsx::addStyle(wb, sheet = sheet, style = general_bold, rows = rowbold, cols = colbold, gridExpand = TRUE, stack = TRUE)
  openxlsx::addStyle(wb, sheet = sheet, style = tableheaderformat, rows = rowtablehead, cols = coltablehead, gridExpand = TRUE, stack = TRUE)
  }


import_header(data, sheet= 1, rowbold = 0, colbold = 0)

#inputting and formating the footer
import_footer <- function (data, sheet= currentsheet, rowbold = (header_height+table_height+1):nrow(data), colbold= 1:(ncol(data))){
  footerdata <- data[sum(header_height + table_height+1):nrow(data),]
  writeData(wb, sheet = sheet, footerdata, startRow = sum ( header_height + table_height+1), colNames = FALSE)
  openxlsx::addStyle(wb, sheet = sheet, style = general_bold, rows = rowbold, cols = colbold, gridExpand = TRUE, stack = TRUE)}

import_footer(data, sheet = 1, colbold = 1)

saveWorkbook(wb, "Automation_tool.xlsx", overwrite = TRUE)
  