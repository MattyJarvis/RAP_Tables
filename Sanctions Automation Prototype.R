###PLEASE NOTE: This is currently an early prototype, several developments, such as accessing the Stat Xplore API to loop in data and more suscinct coding will be developed in the future####


library(XLConnect)
library(huxtable)
library(openxlsx)


##publication dates, *NEEDS EDITTING CURRENTLY, SHOULD BE AUTOMATED*
published <- c("15 May 2018")
next_publish <- c("14 August 2018")


##contact info, edit as apporpriate, nb if you remove an address line, make sure to remove it further in the code
statistician <- c("Statistician: Tracy Hills")
statistician_telephone <- c("Telephone: 0191 216 8223")
statistician_email<- c("Email: tracy.hills@dwp.gsi.gov.uk")
addressline1 <- c("BP5201 Dunstanburgh House")
addressline2<- c("Benton Park Road")
addressline3<- c("Newcastle")
addressline4<- c("NE98 1YX")
press_enq <- c("Press Enquiries: 0203 267 5144")

##websinfo
release_website <- c("https://www.gov.uk/government/collections/jobseekers-allowance-sanctions")
supporting_guides <-c("https://www.gov.uk/government/collections/jobseekers-allowance-sanctions")


#drawing items in table of contents, nb placeholder exists until we construct the hyperlinks in the workbook
table_reference <- c("Table", "1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.7a", "1.8", "1.9", "1.9a", "2.1", "2.2", "2.3", "2.4", "2.5", "2.6", "2.7", "2.8", "3.1", "3.2", "3.3", "3.4", "3.5", "3.6", "3.7", "3.8", "4.1", "4.2", "4.3", "4.4", "4.5")
table_description<- c("Table Description", rep("placeholder", 32))

##create the table
table_of_contents <- data.frame(table_reference, table_description)

#noting the height of the table
table_height <- nrow(table_of_contents)

##convert to a 'huxtable'- a tool and format that allow for us to edit and format tables
table_of_contents<- hux(table_of_contents)

##add headers

#Inserting Title and Subtitles
table_of_contents<- rbind(c("",""), table_of_contents)
View(table_of_contents)
table_of_contents<- rbind(c("To return to contents click 'Back to Contents' link at the top of each page",NA), table_of_contents)
table_of_contents<- rbind(c("To access  data tables select the table heading or tabs.",NA), table_of_contents)
table_of_contents<- rbind(c(NA,NA), table_of_contents)
table_of_contents<- rbind(c("Contents",NA), table_of_contents)
table_of_contents<- rbind(c(NA,NA), table_of_contents)
table_of_contents<- rbind(c("Frequency: Quarterly",NA), table_of_contents)
table_of_contents<- rbind(c("Theme: Social and Welfare",NA), table_of_contents)
table_of_contents<- rbind(c("Coverage: Great Britain",NA), table_of_contents)
table_of_contents<- rbind(c("Next Publication:", next_publish), table_of_contents)
table_of_contents<- rbind(c("Published:", published), table_of_contents)
table_of_contents<- rbind(c(NA,NA), table_of_contents)
table_of_contents<- rbind(c("Benefit Sanctions Statistics",NA), table_of_contents)
table_of_contents<- rbind(c(NA,NA), table_of_contents)
table_of_contents<- rbind(c(NA,NA), table_of_contents)
View(table_of_contents)

#this will be used to define the formatting of footer contents, so the removal 
#of an item from the table of contents does not result in incorrect formatting of footers
header_table_row_number<- nrow(table_of_contents) 

##inserting footers
footer_function <- function(footer_number, footer_text) {
  insert_row(table_of_contents, c(footer_number, footer_text), after = nrow(table_of_contents))
}





table_of_contents <- footer_function(NA, NA)
table_of_contents <- footer_function(NA, NA)
table_of_contents <- footer_function("Contacts", NA)
table_of_contents <- footer_function(NA, NA)
table_of_contents <- footer_function(statistician, NA)
table_of_contents <- footer_function(statistician_telephone, NA)
table_of_contents <- footer_function(statistician_email, NA)
table_of_contents <- footer_function(addressline1, NA)
table_of_contents <- footer_function(addressline2, NA)
table_of_contents <- footer_function(addressline3, NA)
table_of_contents <- footer_function(addressline4, NA)
table_of_contents <- footer_function(press_enq, NA)
table_of_contents <- footer_function("We welcome feedback", NA)
table_of_contents <- footer_function(NA, NA)
table_of_contents <- footer_function("Further information", NA)
table_of_contents <- footer_function(NA, NA)
table_of_contents <- footer_function("Supporting guides for this release:", NA)
table_of_contents <- footer_function(supporting_guides, NA)
table_of_contents <- footer_function(NA, NA)
table_of_contents <- footer_function("Website for this release:", NA)
table_of_contents <- footer_function(release_website, NA)


##we now start formatting the document##

#first the "spare" column to the left needs to be added and the column with "official and official experimental statistics
table_of_contents$A <- c(rep(NA, nrow(table_of_contents)))
table_of_contents$B <- c(rep(NA, 3), "Official and Official Experimental Statistics", rep(NA, sum(nrow(table_of_contents)-4)))

#column A needs to be put first
table_of_contents<- table_of_contents[,c("A", "V1", "V2", "B")]
colnames(table_of_contents) <- c(rep("", 4))

#create workbook to write files into
wb<- createWorkbook()

#create worksheets for sanctions table
addWorksheet(wb, "Table of Contents")
addWorksheet(wb, "Guidance Sheet")
addWorksheet(wb, "Table_1_1")
addWorksheet(wb, "Table_1_2")
addWorksheet(wb, "Table_1_3")
addWorksheet(wb, "Table_1_4")
addWorksheet(wb, "Table_1_5")
addWorksheet(wb, "Table_1_6")
addWorksheet(wb, "Table_1_7")
addWorksheet(wb, "Table_1_7a")
addWorksheet(wb, "Table_1_8")
addWorksheet(wb, "Table_1_9")
addWorksheet(wb, "Table_1_9a")
addWorksheet(wb, "Table_1_9a")
addWorksheet(wb, "Table_2_1")
addWorksheet(wb, "Table_2_2")
addWorksheet(wb, "Table_2_3")
addWorksheet(wb, "Table_2_4")
addWorksheet(wb, "Table_2_4")
addWorksheet(wb, "Table_2_5")
addWorksheet(wb, "Table_2_6")
addWorksheet(wb, "Table_2_7")
addWorksheet(wb, "Table_2_8")
addWorksheet(wb, "Table_3_1")
addWorksheet(wb, "Table_3_2")
addWorksheet(wb, "Table_3_3")
addWorksheet(wb, "Table_3_4")
addWorksheet(wb, "Table_3_5")
addWorksheet(wb, "Table_3_6")
addWorksheet(wb, "Table_3_7")
addWorksheet(wb, "Table_3_8")
addWorksheet(wb, "Table_4_1")
addWorksheet(wb, "Table_4_2")
addWorksheet(wb, "Table_4_3")
addWorksheet(wb, "Table_4_4")
addWorksheet(wb, "Table_4_5")

writeData(wb, sheet = 1, table_of_contents)

#create styles to format headers and footers
titleformat <- openxlsx::createStyle(fontName = NULL, fontSize=16, fontColour = NULL, wrapText = FALSE, textDecoration = 'bold')
subtitleformat <- openxlsx::createStyle(fontName = NULL, fontSize = 14, fontColour = NULL, wrapText = FALSE, textDecoration = 'bold')
general_bold <- createStyle(fontName = NULL, fontColour = NULL, wrapText = FALSE, textDecoration = 'bold')
tableheaderformat <- createStyle(fontName = NULL, fontColour = NULL, wrapText = TRUE, textDecoration = 'bold', border = "bottom", borderColour = getOption("openxlsx.borderColour", "black"))
tablecolumnclassifier <- createStyle(fontName = NULL, fontColour = NULL, wrapText = FALSE, textDecoration = 'bold', border = "bottom", borderColour = getOption("openxlsx.borderColour", "black"))




#formatting text/numbers
#Benefit Sanctions
openxlsx::addStyle(wb, sheet = 1, titleformat, rows = 4, cols = 2)
#Official and Experimental statistics
openxlsx::addStyle(wb, sheet = 1, general_bold , rows = 5, cols = 4)
#publication info
openxlsx::addStyle(wb, sheet = 1, general_bold, rows = 6:10, cols = 2)
#contents
openxlsx::addStyle(wb, sheet = 1, subtitleformat, rows = 12, cols= 2 )
#Table headers
openxlsx::addStyle(wb, sheet = 1, tableheaderformat, rows = 17, cols= 2:3 )
#NUMBER FORMATTING
#contacts
openxlsx::addStyle(wb, sheet = 1, subtitleformat, rows = 52, cols = 2)
#statistician
openxlsx::addStyle(wb, sheet = 1, general_bold, rows = 54, cols = 2)
#press enquiries
openxlsx::addStyle(wb, sheet = 1, general_bold, rows = 61, cols = 2)
#further information
openxlsx::addStyle(wb, sheet = 1, titleformat, rows = 64, cols = 2)

##websinfo
release_website <- c("https://www.gov.uk/government/collections/jobseekers-allowance-sanctions")
supporting_guides <-c("https://www.gov.uk/government/collections/jobseekers-allowance-sanctions")

#adding extrnal hyperlinks
x <- c(release_website)
names(x) <- c(release_website)
class(x) <- "hyperlink"
writeData(wb, sheet = 1, x = x, startCol = 2, startRow = 67 )
x <- c(supporting_guides)
names(x) <- c(supporting_guides)
class(x) <- "hyperlink"
writeData(wb, sheet = 1, x = x, startCol = 2, startRow = 70 )

#internal link

writeFormula(wb, sheet = 1, startRow = 18, startCol = 3, x = makeHyperlinkString(sheet = "Guidance Sheet", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 19, startCol = 3, x = makeHyperlinkString(sheet = "Table_1_1", row = 1, col =2, text = "JSA Sanctions - Decisions by Month"))
writeFormula(wb, sheet = 1, startRow = 20, startCol = 3, x = makeHyperlinkString(sheet = "Table_1_2", row = 1, col =2, text = "JSA Sanctions - New Regime: Decisions by Jobcentre Plus Office"))
writeFormula(wb, sheet = 1, startRow = 21, startCol = 3, x = makeHyperlinkString(sheet = "Table_1_3", row = 1, col =2, text = "JSA Sanctions - New Regime: Decisions to apply a sanctions by Jobcentre Plus Office and Month"))
writeFormula(wb, sheet = 1, startRow = 22, startCol = 3, x = makeHyperlinkString(sheet = "Table_1_4", row = 1, col =2, text = "New Regime: Decisions by Jobcentre Plus District and Level"))
writeFormula(wb, sheet = 1, startRow = 23, startCol = 3, x = makeHyperlinkString(sheet = "Table_1_5", row = 1, col =2, text = "New Regime: Decision to apply a sanction by Level, Reason and Month"))
writeFormula(wb, sheet = 1, startRow = 24, startCol = 3, x = makeHyperlinkString(sheet = "Table_1_6", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 25, startCol = 3, x = makeHyperlinkString(sheet = "Table_1_7", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 26, startCol = 3, x = makeHyperlinkString(sheet = "Table_1_7a", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 27, startCol = 3, x = makeHyperlinkString(sheet = "Table_1_8", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 28, startCol = 3, x = makeHyperlinkString(sheet = "Table_1_9", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 29, startCol = 3, x = makeHyperlinkString(sheet = "Table_1_9a", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 30, startCol = 3, x = makeHyperlinkString(sheet = "Table_2_1", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 31, startCol = 3, x = makeHyperlinkString(sheet = "Table_2_2", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 32, startCol = 3, x = makeHyperlinkString(sheet = "Table_2_3", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 33, startCol = 3, x = makeHyperlinkString(sheet = "Table_2_4", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 34, startCol = 3, x = makeHyperlinkString(sheet = "Table_2_5", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 35, startCol = 3, x = makeHyperlinkString(sheet = "Table_2_6", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 36, startCol = 3, x = makeHyperlinkString(sheet = "Table_2_7", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 37, startCol = 3, x = makeHyperlinkString(sheet = "Table_2_8", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 38, startCol = 3, x = makeHyperlinkString(sheet = "Table_3_1", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 39, startCol = 3, x = makeHyperlinkString(sheet = "Table_3_2", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 40, startCol = 3, x = makeHyperlinkString(sheet = "Table_3_3", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 41, startCol = 3, x = makeHyperlinkString(sheet = "Table_3_4", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 42, startCol = 3, x = makeHyperlinkString(sheet = "Table_3_5", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 43, startCol = 3, x = makeHyperlinkString(sheet = "Table_3_6", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 44, startCol = 3, x = makeHyperlinkString(sheet = "Table_3_7", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 45, startCol = 3, x = makeHyperlinkString(sheet = "Table_3_8", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 46, startCol = 3, x = makeHyperlinkString(sheet = "Table_4_1", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 47, startCol = 3, x = makeHyperlinkString(sheet = "Table_4_2", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 48, startCol = 3, x = makeHyperlinkString(sheet = "Table_4_3", row = 1, col =2, text = "Guidance Sheet"))
writeFormula(wb, sheet = 1, startRow = 49, startCol = 3, x = makeHyperlinkString(sheet = "Table_4_5", row = 1, col =2, text = "Guidance Sheet"))

#setting column widths
setColWidths(wb, sheet =1, cols= 2, widths = 14.2)
setColWidths(wb, sheet = 1, cols = 3, widths = 80)

#inserting dwp logo NB CURRENTLY USING PLACEHOLDER
insertImage(wb, file = "DWPLogo.jpg", sheet = 1, width =3, height =.3, startRow =2, startCol = 2)


########################GUIDANCE SHEET########################################

addWorksheet(wb, "Guidance Sheet")
########################SHEET_1_1##############################################



#read in data
JSA<-read.csv("~/Sanctions Automation/Raw_Data/JSA.csv", skip = 10, colClasses = c(rep("character", 11)),  stringsAsFactors = FALSE)



#filter out columns and rows we don't want
JSArows <- nrow(JSA)
JSA <- JSA[-c(sum(JSArows -18):JSArows),-c(3, 5, 7, 9:11)]

#set decimel places
options(digits = 0)

##firstly we need to isolate the dates in the strings and identify them as dates
JSA$Year <- substring(JSA[,1], 1, 4)
JSA[,1] <- substring(JSA[,1], 5, 6)

#converting column one to month names
months_vector <- c("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
JSA[,1] <- as.numeric(as.character(JSA[,1]))
JSA[,1] <- month.name[JSA[,1]]

#getting month and year for title (added later)
title_date <- paste0(JSA[nrow(JSA), 1], " ", JSA[nrow(JSA), 6])

#items need to be identified as numeric

JSA[,2] <- as.numeric(as.character(JSA[,2]))
JSA[,3] <- as.numeric(as.character(JSA[,3]))
JSA[,4] <- as.numeric(as.character(JSA[,4]))
JSA[,5] <- as.numeric(as.character(JSA[,5]))
JSA[,6] <- as.numeric(as.character(JSA[,6]))

#filtering year column so that only the first label of each year appears
JSA_year_organising<- duplicated(JSA$Year)
View(JSA_year_organising)
for (i in 1:nrow(JSA)){
  if(JSA_year_organising[i] == TRUE){
    JSA[i, "Year"] <- paste0(" ")} }
View(JSA)



##Insert 'Total' line 
JSA <- insert_row(JSA, c("Total", sum(JSA[, 2]), sum(JSA[, 3]), sum(JSA[, 4]), sum(JSA[, 5])), after = nrow(JSA))

JSA <- JSA[-c(nrow(JSA)),]


#name columns and order them 
colnames(JSA) <- c("Month\U00B2", "Decision to apply a sanction\U00B3", "Decision to not apply a sanction (non-adverse)\U2074", "Reserved Decisions\U2075", "Cancelled Referrals\U2076", "Year")
JSA <- JSA[c("Year", "Month\U00B2", "Decision to apply a sanction\U00B3", "Decision to not apply a sanction (non-adverse)\U2074", "Reserved Decisions\U2075", "Cancelled Referrals\U2076")]


#this function means that only the first instance of each year appears on the tabel
for (Year in JSA){
  
  if( unique(JSA[,"Year"]) == FALSE){
    
    JSA[,Year] <- c("")}}


#Number of rows in JSA table, used for formatting later
JSA_table_rows <- nrow(JSA)

#Inserting Title and Subtitle

title <- paste0("April 2000 to ", title_date)
JSA <- insert_row(JSA, c(rep("", 3), "Other decisions taken:", rep(NA, 2)))
JSA <- insert_row(JSA, c(rep("", 6)))
JSA <- insert_row(JSA, c(title, rep(NA, 5)))
JSA <- insert_row(JSA,c("1.1 Jobseekers Allowance Sanction Decisions by Month", rep(NA, 5)))
JSA <- insert_row(JSA, c(rep("", 6)))
JSA <- insert_row(JSA, c('Back to Contents', rep(NA, 5)))
JSA <- insert_row(JSA,c(rep("", 6)))

# #number of rows in header
JSA_header_rows <- nrow(JSA) - JSA_table_rows

#inserting footers

footer_function <- function(footer_number, footer_text) {
  insert_row(JSA, c(footer_number, footer_text, rep(NA, 4)), after = nrow(JSA))
}


JSA <- footer_function("", "")
JSA <- JSA[-c(nrow(JSA)), ]
JSA <- footer_function("Source:", "Decision Making and Appeals System (DMAS), via Stat-Xplore")
JSA <- JSA[-c(nrow(JSA)), ]
JSA <- footer_function("", "")
JSA <- JSA[-c(nrow(JSA)), ]
JSA <- insert_row(JSA, c("Notes:", rep("", 5)), after = nrow(JSA))
JSA <- JSA[-c(nrow(JSA)), ]
JSA <- footer_function(1, "Cells in this Table_have had statistical disclosure control applied to avoid the release of confidential data.  Due to adjustments totals may not be the sum of the individual cells.")
JSA <- JSA[-c(nrow(JSA)), ]
JSA <- footer_function(2, "Month of decision uses the date of decision to allocate to a time period.  If a case was referred in one month but not decided until the next,")
JSA <- JSA[-c(nrow(JSA)), ]
JSA <- footer_function("", "then it will be counted in the decision month.")
JSA <- JSA[-c(nrow(JSA)), ]
JSA <- footer_function(3, "A decision found against the claimant, i.e. a sanction to be applied or the JSA claim is closed (disallowance).")
JSA <- JSA[-c(nrow(JSA)), ]
JSA <- footer_function(4, "A decision found in favour of the claimant, i.e. a sanction or disallowance is not applied.")
JSA <- JSA[-c(nrow(JSA)), ]
JSA <- footer_function(5, "A case would be re-referred if the claimant reclaims JSA within the period of the reserved decision.")
JSA <- JSA[-c(nrow(JSA)), ]
JSA <- footer_function("", "A case would be re-referred if the claimant reclaims JSA within the period of the reserved decision.")
JSA <- JSA[-c(nrow(JSA)), ]
JSA <- footer_function(6, "A cancelled referral results in no sanction decision being made.  This can occur in specific circumstances, for example, the sanction referral")
JSA <- JSA[-c(nrow(JSA)), ]
JSA <- footer_function("", "has been made in error, the claimant stops claiming before they actually committed the sanctionable failure, or information requested by the decision maker.")
JSA <- JSA[-c(nrow(JSA)), ]
JSA <- footer_function("", "was not made available within a specified time period.")
JSA <- JSA[-c(nrow(JSA)), ]

#number of rows in footer
JSA_footer_rows <- sum(nrow(JSA) - JSA_header_rows - JSA_table_rows) 

#adding blank column for formatting
JSA <- insert_column(JSA, c(rep(NA, nrow(JSA))), after = 3)
JSA <- JSA[,c(-5)]
View(JSA)

#giving blank column a name
colnames(JSA) <- c("Year", "Month\U00B2", "Decision to apply a sanction\U00B3", "", "Decision to not apply a sanction (non-adverse)\U2074",  "Reserved Decisions\U2075", "Cancelled Referrals\U2076")
View(JSA)

#we need to convert Sheet_1_1 into a workbook 

JSAbody <- JSA[JSA_header_rows:sum(JSA_header_rows+JSA_table_rows), 1:7]
JSAheader <-JSA[1:JSA_header_rows, 1:7]
JSAfooter <-JSA[sum(JSA_header_rows+JSA_table_rows):sum(JSA_header_rows+JSA_table_rows+JSA_footer_rows), 1:7]

JSAbody[,1] <- as.numeric(JSAbody[,1], na.rm = TRUE)

JSAbody[,3] <- as.numeric(JSAbody[,3])
JSAbody[,4] <- as.numeric(JSAbody[,4])
JSAbody[,5] <- as.numeric(JSAbody[,5])
JSAbody[,6] <- as.numeric(JSAbody[,6])
JSAbody[,7] <- as.numeric(JSAbody[,7])

# for (i in JSAbody[,1]) {
#  if (is.na(JSAbody[i,1]) == FALSE ){
#    JSAbody<- insert_row(JSAbody, c(rep(NA, 6)), after = sum(i - 1))
#    i <- i+1}
#    else {next}}
# 


noborder<- openxlsx::createStyle(fontName = NULL, fontSize=NULL, fontColour = NULL, wrapText = FALSE, border = "topbottomleftright" , borderColour = "white")
openxlsx::addStyle(wb, sheet = 3, style = noborder, rows = 1:10000, cols = 1:20, gridExpand = TRUE, stack = TRUE)

#write data into Table_1_1
openxlsx::writeData(wb, sheet = 3, x = JSAbody, startCol = 2, startRow = JSA_header_rows+2, borders = "all", borderColour = getOption("openxlsx.borderColour", "white"))
View(JSAbody)
#create a style to make numbers have commas and apply it to table
numwithcomma <- openxlsx::createStyle(fontName = NULL, fontSize=NULL, fontColour = NULL, numFmt= "COMMA", wrapText = FALSE)
openxlsx::addStyle(wb, sheet = 3, style = numwithcomma, rows = 1:sum(JSA_header_rows +2 + JSA_table_rows +2), cols = 3:8, gridExpand = TRUE, stack = TRUE)

#rename columns so function doesn't reimport column names from JSA pieces into odd places
colnames(JSAheader) <- c(rep("",7))
colnames(JSAfooter) <- c(rep("",7))


##add in the headers and footers to table 1_1
#header
openxlsx::writeData(wb, sheet = 3, x = JSAheader, startCol = 2, startRow = 1, borders = "all", borderColour = getOption("openxlsx.borderColour", "white"))
#footer
openxlsx::writeData(wb, sheet = 3, x =JSAfooter, startCol = 2, startRow = sum(JSA_header_rows+JSA_table_rows+3), borders = "all", borderColour = getOption("openxlsx.borderColour", "white"))


#title formatting
openxlsx::addStyle(wb, sheet = 3, style = titleformat, rows = 5:6, cols = 2, gridExpand = TRUE, stack = TRUE)
#adding line to top of table
openxlsx::addStyle(wb, sheet = 3, style = tableheaderformat, rows = 7, cols = 2:8, stack = FALSE)
#adding format to other decisions taken
openxlsx::addStyle(wb, sheet = 3, style = tablecolumnclassifier, rows = 8, cols = 6:8, stack = FALSE)
#adding formatting to total line
openxlsx::addStyle(wb, sheet = 3, style = tableheaderformat, rows = sum(JSA_header_rows +JSA_table_rows+4), cols = 2:8, stack = FALSE)
#adding formatting to column headers
openxlsx::addStyle(wb, sheet = 3, style = tableheaderformat, rows = sum(JSA_header_rows+2), cols = 2:8, stack = FALSE)
#formatting column 2 to make bold
openxlsx::addStyle(wb, sheet = 3, style = general_bold, rows = JSA_header_rows:nrow(JSA), cols = 2, stack = TRUE)

#column width setting
setColWidths(wb, sheet = 3, cols = 3, widths = c(9))

openxlsx::saveWorkbook(wb, 'SanctionsOpenxlsxPrototype.xlsx', overwrite = TRUE)

