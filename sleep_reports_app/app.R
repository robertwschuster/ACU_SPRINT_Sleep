#-----------------------------------------------------------------------------------------
#
# Load and analyse sleep tracking data
# Robert Schuster (ACU SPRINT)
# October 2022
#
# This is a Shiny web application. You can run the application by clicking
# the 'Run App' button above.
#
#-----------------------------------------------------------------------------------------

library(shiny)
library(shinythemes)
library(DT)
source("sleep_reports_functions.R")

# UI -------------------------------------------------------------------------------------
ui <- fluidPage(
  theme = shinytheme("spacelab"),
  # Application title
  titlePanel(img(src = "ACU_logo.png", height = 70, width = 200)),
  
  # Sidebar
  sidebarLayout(
    sidebarPanel(
      h3("SPRINT Sleep Reports"),
      
      fileInput("file", 
                "Select the files you want to analyse",
                multiple = T,
                accept = c("text/csv",
                           "text/comma-separated-values,text/plain",
                           ".csv")),
      tags$hr(),
      uiOutput("subjectDropdown"),
      tags$hr(), # horizontal line
      numericInput("rsn", "Recommended sleep per night", value = 8),
      numericInput("nse", "Normal sleep efficiency", value = 85),
      
      # # timezones
      # tags$hr(),
      # h5("Time zone in which data was"),
      # textInput("tzc", "... collected", value = "Enter text..."),
      # textInput("tzd", "... downloaded", value = "Enter text..."),
      # p("E.g.: Australia/Brisbane"),
      # HTML("<p>Click <a href='https://en.wikipedia.org/wiki/List_of_tz_database_time_zones'>here</a> for a list of time zones</p>"),
      
      # create sleep reports
      # fileInput("template", 
      #           "OPTIONAL: Select sleep report template file",
      #           multiple = F,
      #           accept = c(".xlsx")),
      checkboxInput("SQRtemp", label = "Include chart for sleep quality rating in report", value = F),
      downloadButton("downloadData", "Create sleep report")
    ),
    
    # Performance metrics table and graphs of each rep
    mainPanel(
      uiOutput('subName'),
      tableOutput('results'),
      uiOutput('plotTabs')
    )
  )
)

# Server logic ---------------------------------------------------------------------------
server <- function(input, output) {
  # remove default input file size restriction (increase to 30MB)
  options(shiny.maxRequestSize = 30*1024^2)
  # Subject names
  subjects <- reactive({
    if (is.null(input$file)) {
      subjects <- ""
    } else {
      # remove numbers and punctuation from filenames
      # subjects <- unique(gsub('[[:digit:]]+', '', sub("([A-Za-z]+_[A-Za-z0-9]+).*", "\\1", basename(input$file$name))))
      subjects <- unique(sub("^([^_]*_[^_]*).*", "\\1", basename(input$file$name)))
      for (s in 1:length(subjects)) {
        if (grepl("_",substr(subjects[s],nchar(subjects[s]),nchar(subjects[s])))) {
          subjects[s] <- gsub('[[:punct:]]', '', subjects[s])
        }
      }
      subjects <- as.list(subjects)
    }
    return(subjects)
  })
  
  # Subject selection
  output$subjectDropdown <- renderUI({
    selectInput("subSelect", "Select subject", choices = subjects())
  })
  
  # Load files into workspace
  getData <- reactive({
    if (input$subSelect != "" && !is.null(input$subSelect)) {
      i <- grep(input$subSelect, input$file$name)
      data <- importFiles(input$file$datapath[i],input$file$name[i]) # import and prepare files
      data <- sleepData(data, input$rsn, input$nse) # calculate sleep metrics
      return(data)
    }
  })
  
  # Summary table and plots
  output$subName <- renderPrint(h3(input$subSelect))
  output$results <- renderTable(getData()$dex[,1:3], rownames = T)
  output$plotTabs <- renderUI({
    if (input$subSelect != "" && !is.null(input$subSelect)) {
      tabsetPanel(type = "tabs",
                  tabPanel("Duration", renderPlot({
                    durationPlot(getData()$slf, input$rsn)
                  })),
                  tabPanel("Efficiency", renderPlot({
                    efficiencyPlot(getData()$slf, input$nse)
                  })),
                  tabPanel("Pattern", renderPlot({
                    patternPlot(getData()$slf, getData()$sld)
                  }))
      )
    }
  })
  
  # Report generation
  output$downloadData <- downloadHandler(
    filename = function() {
      paste0(input$subSelect,'_sleep-report.xlsx')
    },
    content = function(file) {
      if (input$SQRtemp) {
        temp <- './www/Report template-SQR.xlsx'
      } else {
        temp <- './www/Report template.xlsx'
      }
      saveWorkbook(sleepReport(getData(), temp), file = file)
    }
  )
}

# Run the application --------------------------------------------------------------------
shinyApp(ui = ui, server = server)
