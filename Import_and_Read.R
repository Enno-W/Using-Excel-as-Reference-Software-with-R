####Exporting References from Excel to your .bib - file####
###if you have a .qmd or .rmd file that you work with, you can add this as a chunk right after the YAML-header. Like this, The .bib-file will update every time you render. 
#|echo: false
#Install and load the xlsx package if not already installed
if (!requireNamespace("xlsx", quietly = TRUE)) {
  install.packages("xlsx")
}
library(xlsx) # You might need to install JAVA: https://www.java.com/en/download/manual.jsp
# importing data from sheet 2, no headers. This will lead to conflicts if you have the file open. You can edit your Excel file online though, it will read the data just fine without corrupting the file, e.g. with onedrive. It takes a minute to update the file on disk from Onedrive. 
excel_data <- read.xlsx("Excel_References.xlsx", sheetIndex = 2, header = FALSE)

# Extract citation keys
citation_keys <- as.character(excel_data[, 1])

# Write citation keys to references.bib
writeLines(citation_keys, "bibliography.bib")#You might need to change this to the name of your bibliography

####Importing references from a bibtex-file####
#You can keep this seperate or outcommented from your .qmd or .rmd - file and run it only if necessary
#Install "openxls"
if (!requireNamespace("openxlsx", quietly = TRUE)) {
  install.packages("openxlsx")
}
library(openxlsx)
# read the .bib - file in the destination "Citation_Import/export.bib"
bib_file <- readLines("Citation_Import/export.bib", warn = FALSE)# This works best if you export a BibLaTeX File from Zotero. 

# create the text as one chunk
text <- paste(bib_file, collapse = " ")

# seperate the text in individual entries
entries <- unlist(strsplit(text, split = "\\@"))

# remove empty entries
entries <- entries[entries != ""]

# add "@" to every entry again
entries <- paste("@", entries, sep = "")

# create a dataframe with the entries as a column
data <- data.frame(entries)

# load the workbook
wb <- openxlsx::loadWorkbook("LiteratureAPA.xlsx")#YOUR FILE NAME HERE

# Wähle das Blatt aus, in das du die Daten schreiben möchtest (z.B. "Sheet1")
writeData(wb, "Import", data, colNames = FALSE)

# Speichere die Arbeitsmappe
openxlsx::saveWorkbook(wb, "LiteratureAPA.xlsx", overwrite = TRUE)
