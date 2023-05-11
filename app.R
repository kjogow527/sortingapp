#
# This is a Shiny web application. You can run the application by clicking
# the 'Run App' button above.
#
# Find out more about building applications with Shiny here:
#
#    http://shiny.rstudio.com/
#


# loading all library's
library(shiny)
library(DT)
library(stringi)
library(XML)
library(readxl)
library(bslib)


thematic::thematic_shiny(font = "auto")

output1 <- data.frame()
status <- "waiting, Download unavailable "
listopt <- c("Name", "Community" , "Project" , "Description" , "Modified", "Round", "Activity", "Department", "Type of engagement", "Modified By")
listoptsec<-c("None", "Name", "Community" , "Project", "Description" , "Modified", "Round", "Activity", "Department", "Type of engagement", "Modified By")



# Define UI for application that draws a histogram
ui <- fluidPage(
 
  theme = bs_theme(bootswatch = "solar"),


  
    # Application title
    titlePanel("Sharepoint communication log sorting"),
    
    fluidRow(
    column(4,
            
    
    # Input: Select a file ----
    fileInput("file1", "Choose .xlsx File",
              multiple = TRUE,
              accept = c(".xlsx")),
    
    #"text/csv", "text/comma-separated-values,text/plain",".csv"
    
    # input 1, Choice / text input_______________________________________
    
    selectInput("choice1", "Filter by: (round 1)", listopt),
    textInput("keyword", "Enter Keyword (round 1)", value = ""),
    
    # input 2, Choice / text input_______________________________________
    
    selectInput("choice2", "Filter by: (round 2)", listoptsec),
    textInput("keyword1", "Enter Keyword (round 2)", value = ""),
    
    # input 3, Choice / text input_______________________________________
    
    selectInput("choice3", "Filter by: (round 3)", listoptsec),
    textInput("keyword2", "Enter Keyword (round 3)", value = ""),
    
   # status text
   
   textOutput("selected_var"),
    
    
    # action button
    
    actionButton("button", "Process Dataset"),
   
   
    
    
  # datatable(output)
  
  # download button
  
  downloadButton("downloadData","Download"),
  
  
  # MH image here
  
 #  img(src = "mh.jpg", height = 90, width = 275),

    ),
  column(8,
         p("Instructions:", style = "font-size:50px"),
         p("1. Please upload a xlsx file produced by the sharepoint data dump", style = "font-size:17px"),
         p("2. Select category you wish to sort for", style = "font-size:17px"),
         p("3. Enter a keyword you wish to search for. Example, if you are searching for Kyle Wog, simply putting kyle or wog in is enough", style = "font-size:17px"),
         p("4. If you wish to sort by more categories, enter details in round 2 and 3. If you do not want to sort a second or third round, leave sort by as none", style = "font-size:17px"),
         p("5. Press process dataset, and wait for text to show complete before downloading dataset", style = "font-size:17px"),
         p("5. After dataset is processed, uploaded file is deleted. If you want to process again either re upload or refresh page", style = "font-size:17px"),
         p("Did you know?", style = "font-size:50px"),
         textOutput("fact"),
         p("", style = "font-size:50px"),
         p("", style = "font-size:50px"),
         p("", style = "font-size:50px"),
         
         DT::dataTableOutput('tbl_b')
         )
    )


  
)



# Define server logic required to draw a histogram
server <- function(input, output) {
  
  source("backend.R")

  
  
  
  
  output$fact <- renderText({ 
  randomfact()
  })
  
  
  
  
  
  
  output$selected_var <- renderText({ 
    status
  })
 

  

  
  
  
    output$distPlot <- renderPlot({
        # generate bins based on input$bins from ui.R
        x    <- faithful[, 2]
        bins <- seq(min(x), max(x), length.out = input$bins + 1)

        # draw the histogram with the specified number of bins
        hist(x, breaks = bins, col = 'darkgray', border = 'white')
       
    })
    

    
  observeEvent(input$button, {
    
      # end of reading csv
    
    status <- "Complete, Ready to download!"
   
    output$selected_var <- renderText({ 
      status
    })
  
  df<-  (input$file1$datapath) 
  
  
  output1 <- functioncompute( input$choice1, input$keyword, df)
  
  #df <- output
  
  if(input$choice2 != "None"){
    print("test")
    print(input$keyword1)
    print(input$choice2)
    
    output1 <- functioncompute1( input$choice2, input$keyword1, output1)
    
  }
  
  if(input$choice3 != "None"){
    
    output1 <- functioncompute1( input$choice3, input$keyword2, output1)
    
  
  }
  #_______________________________________________Download
  
  #__________________________________________________Download end

  
  
  print(output1)
  print("done")
 #print(functioncompute( input$choice1, input$keyword, df))
 
print(output1)
 


# downloader__________________________________________________________________________________________

output$downloadData <- downloadHandler(
  
  filename = function() {"output.csv"},
  
  content = function(file) {
    write.csv(output1, file)
  }
)


output$tbl_b = DT::renderDataTable(output1)
# end of downloader_______________________________________________________________________________________

  
  })

    
     
    
}

# Run the application 
shinyApp(ui = ui, server = server)
