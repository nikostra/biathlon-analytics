library(shiny)

locations = c(
  "antholz",
  "arber",
  "brezno",
  "hochfilzen",
  "nove mesto",
  "oberhof",
  "pokljuka",
  "ridnaun",
  "ruhpolding"
)
disciplines = c("individual", "mass start", "pursuit", "sprint")
load("athletes_data.RData") # load athletes data

shinyUI(fluidPage(
  # Application title
  titlePanel("Predict Biathlon Shooting"),
  
  sidebarLayout(
    sidebarPanel(
      selectizeInput("athlete", "Athlete:", athletes_data$name , selected = "boe johannes thingnes"),
      
      selectInput("location", "Location:", locations),
      
      selectInput("discipline", "Discipline:", disciplines),
      
      selectInput("shooting_nr", "Shooting Number:", c("1", "2", "3", "4")),
      
      actionButton("button", "Predict shooting", class = "btn-primary")
      
    ),
    mainPanel(

      tags$head(tags$style(
        HTML(
          "
        .image-container {
          display: flex;
          flex-wrap: nowrap;
          margin-right: -10px; /* Adjust the margin as needed */
        }
        .image-container > div {
          flex: 0 0 auto;
          padding-right: 0px; /* Adjust the padding as needed */
        }
      "
        )
      )),
      div(
        h3("Prediction results:"),
        strong(textOutput("probs_title")),
        textOutput("probs"),
        p(),
        strong(textOutput("simulate"))
        
      ),
      
      div(
        class = "image-container",
        div(imageOutput("shot1")),
        div(imageOutput("shot2")),
        div(imageOutput("shot3")),
        div(imageOutput("shot4")),
        div(imageOutput("shot5")),
      ),
      div(

      )
      

    )
  )
))
