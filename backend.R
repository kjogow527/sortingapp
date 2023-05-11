library(tidyverse)
library(stringi)
library(readxl)




functioncompute <- function(section, keyword, df){

  
links <- linkextraction(df)  
 # mydata <- read_csv(df)
  mydata <- read_xlsx(df)
  output <- data.frame(mydata)
  
 mydata$links <- c(links[["link"]])
  
  # Data installed in r and ready for use
  #_________________________________Data sorting filter__________________________________________
  
  #Keyword you are looking for
  
  keyword <- as.character(keyword)
  
  #section of choice
  
  section <- as.character(section)
  
  
 # stri_enc_toutf8(keyword, is_unknown_8bit = FALSE, validate = FALSE)
 # stri_enc_toutf8(section, is_unknown_8bit = FALSE, validate = FALSE)
  
  print(keyword)
  print(section)
  
  
  
  #_______________________________End of sorting filter_________________________________________
  
  

 

  
  
  output <- mydata[mydata$Name == "ad32432awd4wad324324324324324324324324324",]
  
  
  
  
  size <- (nrow(mydata))
  
  #_________________________________________Primary sorting_____________________________________
  for(i in 1:size){
    
    condition <- grepl((toupper(as.character(keyword))), toupper((mydata[[as.character(section)]][i])), fixed=TRUE)
    
    
    if(is.na(condition)){
      condition = FALSE
    }
    
    
    if((condition) == TRUE){
      output <- rbind(output,mydata[i,])
      
    }
    
  }
  
  output <- subset(output, select = -c(Modified, From, To, Path, Department, `Checked Out To`, `Item Type`,`Modified By`))
  
  

  
  #Name, Community, Project, Received, Description, Modified, Round, Activity, Department, `Type of engagement`, `Modified By`
  write.csv(output,"C:\\Users\\kwog\\Desktop\\output.csv", row.names = FALSE)
  write.csv(output,df, row.names = FALSE)
  return(output)
 
  mydata$`Item Type`
}


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




randomfact <- function(){
 
  listfacts <- list("Manitoba Hydro is is the biggest debt holder in the province with $23.5 Billion of debt.","Manitoba Hydro place is the fifth most energy efficent building in the world.","More than 98% of electricity generated in Manitoba comes from clean and renewable sources such as hydroelectricity and wind.","The high rise elevators at Manitoba Hydro place operate at less than 1600fpm.","In a three phase electrical circut the sin waves are 120 degrees apart. ","A HVDC valve converts AC to DC by seperating the positive and negetive potentials in a three phase circuit.", "Back when dinosaurs existed, there used to be volcanoes that were erupting on the moon.","75% of the world’s diet is produced from just 12 plant and five different animal species.","At birth, a baby panda is smaller than a mouse.","There are less than 30 ships in the Royal Canadian Navy which are less than most third-world countries.","Johnny Cash took only three voice lessons before his teacher advised him to stop taking lessons and to never deviate from his natural voice.","Violin bows are commonly made from horse hair.","Lettuce is a member of the sunflower family.","The average American child is given $3.70 per tooth that falls out.","Vanilla flavoring is sometimes made with the urine of beavers.","If you heat up a magnet, it will lose its magnetism.","More tornadoes occur in the United Kingdom per square mile than any other country in the world.","Tic Tacs got their name from the sound they make when they are tossed around in their container.","Costa Coffee employs Gennaro Pelliccia as a coffee taster, who has had his tongue insured for £10 million since 2009.","The color red doesn’t really make bulls angry, they are color-blind.","IKEA is an acronym which stands for Ingvar Kamprad Elmtaryd Agunnaryd, which is the founder’s name, farm where he grew up, and hometown.","Although GPS is free for the world to use, it costs $2 million per day to operate. The money comes from American tax revenue.","Sour Patch Kids are from the same manufacturer as Swedish Fish. The red Sour Patch Kids are the same candy as Swedish Fish, but with sour sugar.","In Colorado, USA, there is still an active volcano. It last erupted about the same time as the pyramids were being built in Egypt.","The only letter that doesn’t appear on the periodic table is J.","The Lego Group is the world’s most powerful brand. There are more Lego Minifigures than there are people on Earth.","The largest known prime number has 17,425,170 digits. The new prime number is 2 multiplied by itself 57,885,161 times, minus 1.")
  
 number <- runif(1, min = 1, max = 27) 
 number <- round(number)
 
string <- as.character(listfacts[number])
 
print(string)
 
 return(string)
  
  
  
  
}





themecol <- function(){

color <- c("minty","solar","vapor","slate","sketchey")
value <- runif(1, min = 1, max = 5) 
value <- round(value)
themeval <- (color[value])
print(themeval)
return(themeval)

}   















functioncompute1 <- function(section, keyword, df){
  
  
  mydata <- data.frame(df)
  
  
  output <- data.frame(mydata)
  
  
  
  # Data installed in r and ready for use
  #_________________________________Data sorting filter__________________________________________
  
  #Keyword you are looking for
  
  keyword <- as.character(keyword)
  
  #section of choice
  
  section <- as.character(section)
  
  
  # stri_enc_toutf8(keyword, is_unknown_8bit = FALSE, validate = FALSE)
  # stri_enc_toutf8(section, is_unknown_8bit = FALSE, validate = FALSE)
  
  print(keyword)
  print(section)
  
  
  
  #_______________________________End of sorting filter_________________________________________
  
  
  
  
  
  
  
  output <- mydata[mydata$Name == "ad32432awd4wad324324324324324324324324324",]
  
 
  
  size <- (nrow(mydata))
  
  #_________________________________________Primary sorting_____________________________________
  for(i in 1:size){
    
    condition <- grepl((toupper(as.character(keyword))), (toupper(mydata[[as.character(section)]][i])), fixed=TRUE)
    
    
    if(is.na(condition)){
      condition = FALSE
    }
    
    
    if((condition) == TRUE){
      output <- rbind(output,mydata[i,])
      
    }
    
  }
  
 # output <- subset(output, select = -c(Received, Modified, From, To, Path, Department, `Checked Out To`, `Item Type`))
  
  #Name, Community, Project, Received, Description, Modified, Round, Activity, Department, `Type of engagement`, `Modified By`
  write.csv(output,"c:\\users\\kwog\\desktop\\output.csv", row.names = FALSE)
  #write.csv(output,df, row.names = FALSE)
  return(output)
  
  mydata$`Item Type`
}














#___________________________________________________________________________________________________________________________________________________
