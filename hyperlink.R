library(XML)


output <- linkextraction("Sample data extract from Community Engagement sharepoint.xlsx")
print(length(output))

#linkextraction <- function(my.excel.file){

reference <- 50
links <- list()
path <- list()

#my.excel.file <- "C:\\Users\\kwog\\Desktop\\Sample data extract from Community Engagement sharepoint.xlsx"

# rename file to .zip
my.zip.file <- sub("xlsx", "zip", my.excel.file)
file.copy(from = my.excel.file, to = my.zip.file)

# unzip the file
unzip(my.zip.file)

# unzipping produces a bunch of files which we can read using the XML package
# assume sheet1 has our data
xml <- xmlParse("xl/worksheets//_rels/sheet1.xml.rels")
xmlpath <- xmlParse("xl/worksheets//sheet1.xml")
print(xml)
# finally grab the hyperlinks
hyperlinks <- xpathApply(xml, "//x:Relationship/@Target", namespaces="x")
pathfrom <- xpathApply(xmlpath, "//x:hyperlink/@ref", namepaces="x")

#length of the list for extraction
length <- length(hyperlinks)


for(i in 1:length){
  
  links <- c(links,(hyperlinks[[i]][["Target"]]))
  path <- c(path,(pathfrom[[i]][["ref"]]))
 
}


return(path)
}

write.csv(output)










































print(linkextraction("Sample data extract from Community Engagement sharepoint.xlsx"))

linkextraction <- function(my.excel.file){

links <- list()
path <- list()
linkp <- list()
#my.excel.file <- "C:\\Users\\kwog\\Desktop\\Sample data extract from Community Engagement sharepoint.xlsx"


# rename file to .zip
my.zip.file <- sub("xlsx", "zip", my.excel.file)
file.copy(from = my.excel.file, to = my.zip.file)

# unzip the file
unzip(my.zip.file)

# unzipping produces a bunch of files which we can read using the XML package
# assume sheet1 has our data
xml <- xmlParse("xl/worksheets//_rels/sheet1.xml.rels")
xmlpath <- xmlParse("xl/worksheets//sheet1.xml")
#print(xmlpath)

# finally grab the hyperlinks
hyperlinks <- xpathApply(xml, "//x:Relationship/@Target", namespaces="x")
hyperlinksp <- xpathApply(xml, "//x:Relationship/@Id", namespaces="x")
pathfrom <- xpathApply(xmlpath, "//x:hyperlink/@ref", namespaces="x")


#length of the list for extraction
length <- length(pathfrom)


for(i in 1:length){
  
  links <- c(links,(hyperlinks[[i]][["Target"]]))
  path <- c(path,(as.numeric(gsub('A','',pathfrom[[i]][["ref"]])))-1)
  linkp <- (c(linkp,(gsub('rId', '',hyperlinksp[[i]][["Id"]]))))
  
}

linkdf = data.frame(unlist(links),unlist(as.numeric(linkp)))

names(linkdf) = c("link","ref")



linkdf <-linkdf[order(linkdf$ref),]

linkdf <- subset(linkdf, select = -c(ref))

return(linkdf)
}



