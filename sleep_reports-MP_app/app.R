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
source("sleep_reports-MP_functions.R")

# UI -------------------------------------------------------------------------------------
ui <- fluidPage(
  theme = shinytheme("spacelab"),
  # Application title
  titlePanel(img(src = "ACU_logo.png", height = 70, width = 200)),
  
  # Sidebar
  sidebarLayout(
    sidebarPanel(
      h3("SPRINT Actiwatch Sleep Reports"),
      
      fileInput("file", 
                "Select the files you want to analyse",
                multiple = TRUE,
                accept = c("text/csv",
                           "text/comma-separated-values,text/plain",
                           ".csv")),
      # select subject
      tags$hr(),
      uiOutput("subjectDropdown"),
      # define cycles and phases
      tags$hr(), # horizontal line
      uiOutput("C1"),
      uiOutput("C1P2"),
      uiOutput("C1P3"),
      uiOutput("C2"),
      uiOutput("C2P2"),
      uiOutput("C2P3"),
      
      # create sleep reports
      tags$hr(), # horizontal line
      numericInput("rsn", "Recommended sleep per night", value = 8),
      numericInput("nse", "Normal sleep efficiency", value = 85),
      # fileInput("template", 
      #           "OPTIONAL: Select sleep report template file",
      #           multiple = F,
      #           accept = c(".xlsx")),
      downloadButton("downloadData", "Create sleep report"),
      width = 3
    ),
    
    # Performance metrics table and graphs of each rep
    mainPanel(
      uiOutput("subName"),
      tableOutput("results"),
      uiOutput("plotTabs"),
      width = 9
    )
  )
)

# Server logic ---------------------------------------------------------------------------
server <- function(input, output) {
  # remove default input file size restriction (increase to 30MB)
  options(shiny.maxRequestSize = 30 * 1024^2)
  # Subject names
  subjects <- reactive({
    if (is.null(input$file)) {
      subjects <- ""
    } else {
      # remove numbers and punctuation from filenames
      subjects <- unique(gsub("[[:digit:]]+", "", sub("([A-Za-z]+_[A-Za-z0-9]+).*", "\\1", basename(input$file$name))))
      for (s in 1:length(subjects)) {
        if (grepl("_", substr(subjects[s], nchar(subjects[s]), nchar(subjects[s])))) {
          subjects[s] <- gsub("[[:punct:]]", "", subjects[s])
        }
      }
      subjects <- as.list(subjects)
    }
    return(subjects)
  })
  
  # Subject selection
  output$subjectDropdown <- renderUI({
    selectInput("subSelect", "Select subject", choices = subjects(), multiple = FALSE)
  })
  
  # Load files into workspace
  getData <- reactive({
    if (input$subSelect != "" && !is.null(input$subSelect)) {
      i <- grep(input$subSelect, input$file$name)
      data <- importFiles(input$file$datapath[i], input$file$name[i]) # import and prepare files
      data <- sleepData(data, input$rsn, input$nse) # calculate sleep metrics
      return(data)
    }
  })
  
  # Cycle and phase dates
  output$C1 <- renderUI({
    ifelse(input$subSelect != "" && !is.null(input$subSelect),
           v <- c(as.character(format(as.Date(min(getData()$sld$Date))), "yyyy-mm-dd"),
                  as.character(format(as.Date(max(getData()$sld$Date))), "yyyy-mm-dd")), 
           v <- rep(as.character(format(as.Date(Sys.Date())), "yyyy-mm-dd"), 2))
    dateRangeInput("C1", "Cycle 1", start = v[1], end = v[2])
  })
  
  output$C1P2 <- renderUI({
    v <- ifelse(input$subSelect != "", as.character(format(as.Date(input$C1[1])), "yyyy-mm-dd"),
                as.character(format(as.Date(Sys.Date())), "yyyy-mm-dd"))
    dateInput("C1P2", "Start phase 2", value = v)
  })
  
  output$C1P3 <- renderUI({
    v <- ifelse(input$subSelect != "", as.character(format(as.Date(input$C1P2)), "yyyy-mm-dd"),
                as.character(format(as.Date(Sys.Date())), "yyyy-mm-dd"))
    dateInput("C1P3", "Start phase 3", value = v)
  })
  
  output$C2 <- renderUI({
    ifelse(input$subSelect != "" && !is.null(input$subSelect),
           v <- c(as.character(format(as.Date(input$C1[2])), "yyyy-mm-dd"),
                  as.character(format(as.Date(max(getData()$sld$Date))), "yyyy-mm-dd")), 
           v <- rep(as.character(format(as.Date(Sys.Date())), "yyyy-mm-dd"), 2))
    dateRangeInput("C2", "Cycle 2", start = v[1], end = v[2])
  })
  
  output$C2P2 <- renderUI({
    v <- ifelse(input$subSelect != "", as.character(format(as.Date(input$C2[1])), "yyyy-mm-dd"),
                as.character(format(as.Date(Sys.Date())), "yyyy-mm-dd"))
    dateInput("C2P2", "Start phase 2", value = v)
  })
  
  output$C2P3 <- renderUI({
    v <- ifelse(input$subSelect != "", as.character(format(as.Date(input$C2P2)), "yyyy-mm-dd"),
                as.character(format(as.Date(Sys.Date())), "yyyy-mm-dd"))
    dateInput("C2P3", "Start phase 3", value = v)
  })
  
  # Summary table and plots
  output$subName <- renderPrint(h3(input$subSelect))
  output$results <- renderTable(getData()$dex[, 1:3], rownames = TRUE)
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
      paste0(input$subSelect, "_sleep-report.xlsx")
    },
    content = function(file) {
      temp <- "./www/Report template_MP2.xlsx"
      data <- menstCycles(getData(), input$C1, input$C2, input$C1P2, input$C1P3, input$C2P2, input$C2P3)
      saveWorkbook(sleepReport(data, temp), file = file)
    }
  )
}

# Run the application --------------------------------------------------------------------
shinyApp(ui = ui, server = server)
