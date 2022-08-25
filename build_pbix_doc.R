#### Import Library
library(officer)
library(officedown)
library(magrittr)
library(knitr)
library(pbixr)
library(knitr)
library(RCurl)
library(ggplot2)
library(ggraph)
library(imager)
library(tidyr)
library(formatR)
library(stringr)
library(openxlsx)

# Path to the pbxi file:----

pbix_path_file <- "-------------PATH TO YOUR FILE------------(use two backslash \\ or one forward slash /) "

file_name <- basename(pbix_path_file)

# File Info:----

pbix_file_info <- as.data.frame(file.info(pbix_path_file, extra_cols = TRUE))

# Run and open pbix file: ----
shell.exec(pbix_path_file)

# Close application
### https://stackoverflow.com/questions/55849698/close-external-application-launched-from-r
# system2("taskkill", args = "/im excel.exe")
# system2("taskkill", args = "/im PBIDesktop.exe")


# Metadata ----

## The Contents of the pbix File ----
# File path of the sample
input_pbix <- pbix_path_file
pbi_content <- f_get_pbix_info(input_pbix)
pbi_content_names <- pbi_content$Name

# Identify extensions
pbi_content_types <- unique(tools::file_ext(pbi_content_names))

# Show content
biggest_file <- max(as.numeric(gsub("[^0-9]*","",pbi_content$Length)))
pbi_content$Length <- prettyNum(pbi_content$Length, big.mark = ",")
knitr::kable(pbi_content)


# Find sha256 hash of the file
library(digest)
hash <- digest::digest(pbix_path_file, algo="sha256", file=TRUE)

###_____________________________________________________________________________

# Identify the right port
connections_open <- f_get_connections()
connections_open$pbix <- gsub(" - Power BI Desktop", "",
                              connections_open$pbix_name)
connections_open <- connections_open[which(connections_open$pbix ==
                                             gsub("[.]pbix", "", basename(input_pbix))), ][1, ]
correct_port <- as.numeric(connections_open$ports)
# Construct the connection
connection_db <- paste0("Provider=MSOLAP.8;Data Source=localhost:",
                        correct_port, ";MDX Compatibility=1")

###_____________________________________________________________________________
# Get All Power Query M Code , With Properties ----
m_query_mem <- paste0("SELECT * FROM `$SYSTEM.TMSCHEMA_PARTITIONS")
get_m_query_mem <- f_query_datamodel(m_query_mem, connection_db)

m_query_par <- paste0("SELECT * FROM `$SYSTEM.TMSCHEMA_EXPRESSIONS")
get_m_query_par <- f_query_datamodel(m_query_par, connection_db)


# library(stringr)
# stringr::str_extract(get_m_query_mem$Name, "-")
# stringr::str_split(get_m_query_mem$Name, "-")

get_m_query_mem$Query_Name <- stringr::word(get_m_query_mem$Name,1,sep = "\\-")

get_m_query_mem$Flat_file <- ifelse((stringr::str_detect(get_m_query_mem[,6], "let\n    Source = Csv.Document", negate = FALSE) | stringr::str_detect(get_m_query_mem[,6], "let\n    Source = Excel.Workbook", negate = FALSE)),"Flat file","Other")

stringr::str_match(get_m_query_mem[,6], "let\\n    Source = Excel.Workbook\\(File.Contents\\(\"\\s*(.*?)\\s*\\\"\\), null")
# str_match(a, "STR1\\s*(.*?)\\s*STR2") #https://stackoverflow.com/questions/39086400/extracting-a-string-between-other-two-strings-in-r
Source_path <- stringr::str_match(get_m_query_mem[,6], "let\\n    Source = Excel.Workbook\\(File.Contents\\(\"\\s*(.*?)\\s*\\\"\\), null")
get_m_query_mem$Source_path <- Source_path[,2]

### Select columns for reporting
get_m_query_mem_01 <- get_m_query_mem[,c(2,24,6,8,12,13,25,26)]

###_____________________________________________________________________________
# Get All Model Data Sources ----
d_sources <- paste0("SELECT * FROM `$SYSTEM.DISCOVER_POWERBI_DATASOURCES")
get_d_sources <- f_query_datamodel(d_sources, connection_db)

###_____________________________________________________________________________
# Get All Tables Sources ----
tables <- paste0("SELECT * FROM `$SYSTEM.TMSCHEMA_TABLES")
get_tables <- f_query_datamodel(tables, connection_db)
# get_tables_01 <- get_tables[get_tables$SystemFlags==0, c(1,3,6,12)]
get_tables_01 <- get_tables[, c(1,3,6,8,9, 10,12)]
###_____________________________________________________________________________
# Get All Measures ----
measures <- paste0("SELECT * FROM `$SYSTEM.MDSCHEMA_MEASURES")
get_measures <- f_query_datamodel(measures, connection_db)
get_measures_01 <- get_measures[, c(19,5,14,21,15)]

###_____________________________________________________________________________
# Get All Calculations ----
calculations <- paste0("SELECT * FROM `$SYSTEM.TMSCHEMA_CALCULATION_ITEMS")
get_calculations <- f_query_datamodel(calculations, connection_db)

###_____________________________________________________________________________


# Get relationships ----
sql_relationships <- paste0("select FromTableID, FromColumnID, ToTableID,",
                            " ToColumnID  from `$SYSTEM.TMSCHEMA_RELATIONSHIPS")
get_relationships <- f_query_datamodel(sql_relationships, connection_db)

# Get names of columns
sql_columns <- "select * from `$SYSTEM.TMSCHEMA_COLUMNS"
get_columns <- f_query_datamodel(sql_columns, connection_db) %>%
  mutate(ColumnName = ifelse(nchar(as.character(ExplicitName)) == 0,
                             as.character(InferredName), as.character(ExplicitName))) %>%
  select(ColumnID = ID, ColumnName)

# Get names of tables
sql_table <- "select * from `$SYSTEM.TMSCHEMA_TABLES"
get_tables <- f_query_datamodel(sql_table, connection_db) %>%
  select(TableID = ID, TableName = Name)

# Merge things together
get_relationship_names <- get_relationships %>%
  merge(get_tables, by.x = "FromTableID", by.y = "TableID") %>%
  rename(FromTable = TableName) %>%
  merge(get_tables, by.x = "ToTableID", by.y = "TableID") %>%
  rename(ToTable = TableName) %>%
  merge(get_columns, by.x = "FromColumnID", by.y = "ColumnID") %>%
  rename(FromColumn = ColumnName) %>%
  merge(get_columns, by.x = "ToColumnID", by.y = "ColumnID") %>%
  rename(ToColumn = ColumnName) %>%
  select(FromTable, ToTable, FromColumn, ToColumn)

# Make things all characters
# drop table with long name. Doing this just for display purposes here.
get_relationship_names <- get_relationship_names[-5, ] %>%
  mutate(FromTable = as.character(FromTable)) %>%
  mutate(ToTable = as.character(ToTable))

# Plot the results with igraph
c_list <- list()
for (i in 1:nrow(get_relationship_names)) {
  c_list[[i]] <- get_relationship_names[i, 1:2]
}
ed <- as.character(unlist(c_list))
g1 <- igraph::graph(edges=ed, directed=F)

library(igraph)
set.seed(1492)
plot(g1,
     vertex.color="#E5E5E5", 
     vertex.label.cex = 0.85, 
     vertex.label.degree = -pi/2,
     vertex.size=5, 
     vertex.shape="rectangle", 
     edge.color="red", 
     asp=1)
title("Data Model Used",cex.main=1,col.main="black")

# Export Pic -----------
pic_path <- paste0(getwd(),"/graph.png")
png(filename = pic_path, res=600, width=5000, height=5000, pointsize=10,
    type="windows", antialias="cleartype")
plot(g1,
     vertex.color="#E5E5E5", 
     vertex.label.cex = 0.85, 
     vertex.label.degree = -pi/2,
     vertex.size=5, 
     vertex.shape="rectangle", 
     edge.color="red", 
     asp=1)
title("Data Model Used",cex.main=1,col.main="black")
dev.off()


####____________________________________________________

file.path(pbix_path_file)
basename(pbix_path_file)
dirname(pbix_path_file)
getwd()
dir.create(paste0(getwd(),"/unzip_pbix"))
# Run the function
library(pbixr)
pbixr::f_decompress_pbix(pbix_path_file, paste0(getwd(),"/unzip_pbix"))

###_____________________________________________________________________________

json <- readLines(paste0(getwd(),"/unzip_pbix/Report/Layout"), n = 2, ok = TRUE, warn = TRUE, encoding = "UTF-8", skipNul = TRUE)

json <- jsonlite::parse_json(json, simplifyVector = FALSE)

json_01 <- json

# ## Get all page reports
sections_01 <- json_01$sections

#sapply(sections_01, "[[", 6)

page_reports <- data.frame("Page_reports"= sapply(sections_01, "[[", 2))
page_reports$Hidden_page <- stringr::str_detect(sapply(sections_01, "[[", 6), ",\\\"visibility\"\\:1", negate = FALSE)
page_reports$Filter_on_page <- stringr::str_detect(sapply(sections_01, "[[", 3), "\\[\\]", negate = TRUE)

###_____________________________________________________________________________
### Save Excel File----
##https://stackoverflow.com/questions/17976522/export-both-image-and-data-from-r-to-an-excel-spreadsheet

library(openxlsx)
wb <- openxlsx::createWorkbook()
addWorksheet(wb, "Plots")
addWorksheet(wb, "tables")
addWorksheet(wb, "m_query")
addWorksheet(wb, "measures")
addWorksheet(wb, "page_reports")
writeData(wb, "Plots", "SHA-256",
          startRow = 1,
          startCol = 1,)
writeData(wb, "Plots", "File_path",
          startRow = 2,
          startCol = 1,)
writeData(wb, "Plots", hash,
          startRow = 1,
          startCol = 2,)
writeData(wb, "Plots", input_pbix,
          startRow = 2,
          startCol = 2,)
insertImage(wb, "Plots", pic_path,
            width = 6,
            height = 6,
            startRow = 3,
            startCol = 1,
            units = "in",
            dpi = 600)
writeDataTable(wb, "tables", get_tables_01)
writeDataTable(wb, "m_query", get_m_query_mem_01)
writeDataTable(wb, "measures", get_measures_01)
writeDataTable(wb, "page_reports", page_reports)
openxlsx::saveWorkbook(wb,
                       paste0(getwd(),"/", paste0(sub(pattern = "(.*)\\..*$", replacement = "\\1", basename(pbix_path_file)),"_Documentation_", format(Sys.Date(), "%d.%m.%y"),".xlsx")), overwrite = TRUE)

# file.exists("TestFile.xlsx")

### Call officer_code.R to build Word Document.
source(paste0(getwd(), "/officer_code.R"))

# Close application:----
system2("taskkill", args = "/im PBIDesktop.exe")
