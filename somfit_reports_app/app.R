#-----------------------------------------------------------------------------------------
#
# Load and analyse Somfit sleep tracking data
# Robert Schuster (ACU SPRINT)
# November 2022
#
# This is a Shiny web application. You can run the application by clicking
# the 'Run App' button above.
#
#-----------------------------------------------------------------------------------------

library(shiny)
library(shinythemes)
library(DT)
source("somfit_reports_functions.R")

# UI -------------------------------------------------------------------------------------
ui <- fluidPage(
  theme = shinytheme("spacelab"),
  # Application title
  titlePanel(img(src = "ACU_logo.png", height = 70, width = 200)),
  
  # Sidebar
  sidebarLayout(
    sidebarPanel(
      h3("SPRINT Somfit Sleep Reports"),
      
      fileInput("file", 
                "Select the files you want to analyse",
                multiple = T,
                accept = ".rtf"),
      tags$hr(),
      uiOutput("subjectDropdown"),
      tags$hr(), # horizontal line
      numericInput("rsn", "Recommended sleep per night", value = 8),
      numericInput("nse", "Normal sleep efficiency", value = 85),
      tags$hr(),
      radioButtons("MPCG", label = "Researcher",
                   choices = list("Madi" = 1, "Riss" = 2), selected = 1),
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
      sNames <- ""
    } else {
      # extract subject names from RTF files
      subjects <- subNames(input$file$datapath, input$file$name)
      sNames <- as.list(unique(subjects[,2]))
    }
    return(sNames)
  })
  
  # Subject selection
  output$subjectDropdown <- renderUI({
    selectInput("subSelect", "Select subject", choices = subjects())
  })
  
  # Load files into workspace
  getData <- reactive({
    if (input$subSelect != "" && !is.null(input$subSelect)) {
      i <- grep(input$subSelect, subNames(input$file$datapath,input$file$name)[,2])
      data <- importFiles(input$file$datapath[i],input$file$name[i],input$MPCG) # import and prepare files
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
      paste0(input$subSelect,'_somfit-report.xlsx')
    },
    content = function(file) {
      if (input$MPCG == 1) {
        temp <- './www/Report template_somfit_MP.xlsx'
      } else if (input$MPCG == 2) {
        temp <- './www/Report template_somfit_CG.xlsx'
      }
      saveWorkbook(sleepReport(getData(), temp), file = file)
    }
  )
}

# Run the application --------------------------------------------------------------------
shinyApp(ui = ui, server = server)
