library(officer)
library(officedown)
library(magrittr)
library(knitr)

template_path <- paste0(getwd(),"/","template_for_documentation.docx")
knitr::kable(layout_summary(my_pres))
### Read and import a docx file as an R object representing the document
doc_01 <- read_docx(path = template_path) 
# knitr::kable(layout_summary(doc_01))
doc_01 %>%
  body_add_par(paste0("Documentation for: ", "\'",file_name, "\'")
               , style = "Title") %>%
  body_add_par(value = "Table of content", style = "heading 1") %>% 
  body_add_toc(level = 2) %>%
  body_add_break() %>% 
  body_add_par(value = "File details:", style = "heading 1") %>% 
  body_add_par(value = paste0("File path: ",pbix_path_file), style = "Normal") %>% 
  body_add_par(value = paste0("SHA-256 Digest: ",hash), style = "Normal") %>%
  body_add_par(paste0("Documentation created on the ", format(Sys.Date(), "%d.%m.%y"),". For more details contact:"), style = "Normal") %>%
  body_add_img(src = paste0(getwd(),"/graph.png"),height = 6.5, width = 6.5) %>% 
  body_add_par(value = "Simplified model schema", style = "Image Caption") %>% 
  body_add_par(value = "", style = "Normal") %>% 
  body_add_par(value = "Page Reports:", style = "heading 1") %>% 
  body_add_par(value = "", style = "Normal") %>% 
  body_add_table(page_reports, style = "Table",first_column = TRUE) %>% 
  body_add_par(value = "Tables:", style = "heading 1") %>% 
  body_add_par(value = "", style = "Normal") %>% 
  body_add_table(get_tables_01[get_tables_01$SystemFlags == 0,c(1:5,7)], style = "Table",first_column = TRUE) %>% 
  body_add_par(value = "All system tables are excluded. Full list available in excel file.", style = "Table Caption") %>% 
  body_add_par(value = "M-queries:", style = "heading 1") %>% 
  body_add_par(value = "", style = "Normal") %>% 
  body_add_table(get_m_query_mem_01[get_m_query_mem_01$Type == 4,c(1:2,5:8)], style = "Table",first_column = TRUE) %>% 
  body_add_par(value = "All system tables are excluded. Full list available in excel file.", style = "Table Caption") %>% 
  body_add_par(value = "DAX Code and Measures:", style = "heading 1") %>% 
  body_add_table(get_measures_01[,-3], style = "Table",first_column = TRUE) %>% 
  body_add_par(value = "All system tables are excluded. Full list available in excel file.", style = "Table Caption") %>% 
  body_add_par(value = "", style = "Normal") %>% 
  body_add_par(value = "", style = "Normal") %>% 
  body_add_par(value = "", style = "Normal")



print(doc_01, target = paste0(sub(pattern = "(.*)\\..*$", replacement = "\\1", basename(pbix_path_file)),"_Documentation_", format(Sys.Date(), "%d.%m.%y"),".docx"))
