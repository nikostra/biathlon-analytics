#
# This is the server logic of the shiny web app
#

library(shiny)
library(xgboost)

# load the model, athlete data and a sample model matrix from the provided files
model = xgboost::xgb.load("xgb.model")
load("model_matrix.RData") 
load("athletes_data.RData")

# define some variables
locations = c("antholz","arber","brezno","hochfilzen","nove mesto","oberhof","pokljuka","ridnaun","ruhpolding")
disciplines = c("individual", "mass start", "pursuit", "sprint")
threshold = 0.8072
prediction_loaded = FALSE

# data is model_matrix mit 46 columns
setup_data = function(data){
  #set locations and discipline to NA
  data[,1:46] = NA
  
  #setup shot number
  data[1,20] = 1 
  data[2,20] = 2 
  data[3,20] = 3 
  data[4,20] = 4 
  data[5,20] = 5 
  
  data[,12:14] = 0 # setup season (2022/23)
  data[,15] = 1 
  data[,30:31] = 0 
  data[,29] = 1
  
  return(data)
}


shinyServer(function(input, output) {

  # function triggers if user hits the predict button
  observeEvent(input$button, {
    prediction_loaded = TRUE
    # reset data
    model_matrix = setup_data(model_matrix)
    
    # setup data for new athlete
    athlete_id = which(athletes_data$name == input$athlete)
    model_matrix[,40] = as.numeric(athletes_data$pre_hit_rate_10[athlete_id])
    model_matrix[,42] = as.numeric(athletes_data$pre_hit_rate_50[athlete_id])
    model_matrix[,44] = as.numeric(athletes_data$pre_hit_rate_200[athlete_id])
    
    # set up variables according to shooting mode (shootings 1 and 3 are prone, others are standing)
    if (as.numeric(input$shooting_nr) %% 2 == 0){
      model_matrix[1,23] = as.numeric(athletes_data$shooting_time_s_shot1[athlete_id])
      model_matrix[2:5,23] = as.numeric(athletes_data$shooting_time_s[athlete_id])
      
      model_matrix[,41] = as.numeric(athletes_data$pre_hit_rate_10_mode_s[athlete_id])
      model_matrix[,43] = as.numeric(athletes_data$pre_hit_rate_50_mode_s[athlete_id])
      model_matrix[,45] = as.numeric(athletes_data$pre_hit_rate_200_mode_s[athlete_id])
      
      model_matrix[1,46] = as.numeric(athletes_data$pre_hit_rate_200_mode_s_shotNr_1[athlete_id])
      model_matrix[2,46] = as.numeric(athletes_data$pre_hit_rate_200_mode_s_shotNr_2[athlete_id])
      model_matrix[3,46] = as.numeric(athletes_data$pre_hit_rate_200_mode_s_shotNr_3[athlete_id])
      model_matrix[4,46] = as.numeric(athletes_data$pre_hit_rate_200_mode_s_shotNr_4[athlete_id])
      model_matrix[5,46] = as.numeric(athletes_data$pre_hit_rate_200_mode_s_shotNr_5[athlete_id])
    } else {
      model_matrix[1,23] = as.numeric(athletes_data$shooting_time_p_shot1[athlete_id])
      model_matrix[2:5,23] = as.numeric(athletes_data$shooting_time_p[athlete_id])
      
      model_matrix[,41] = as.numeric(athletes_data$pre_hit_rate_10_mode_p[athlete_id])
      model_matrix[,43] = as.numeric(athletes_data$pre_hit_rate_50_mode_p[athlete_id])
      model_matrix[,45] = as.numeric(athletes_data$pre_hit_rate_200_mode_p[athlete_id])
      
      model_matrix[1,46] = as.numeric(athletes_data$pre_hit_rate_200_mode_p_shotNr_1[athlete_id])
      model_matrix[2,46] = as.numeric(athletes_data$pre_hit_rate_200_mode_p_shotNr_2[athlete_id])
      model_matrix[3,46] = as.numeric(athletes_data$pre_hit_rate_200_mode_p_shotNr_3[athlete_id])
      model_matrix[4,46] = as.numeric(athletes_data$pre_hit_rate_200_mode_p_shotNr_4[athlete_id])
      model_matrix[5,46] = as.numeric(athletes_data$pre_hit_rate_200_mode_p_shotNr_5[athlete_id])
      
    }

      
    # setup location
    model_matrix[,1:8] = 0
    if(input$location != "antholz"){
      model_matrix[,which(grepl(input$location, colnames(model_matrix)))] = 1
    }
    
    # setup discipline
    model_matrix[,9:11] = 0
    if(input$discipline != "individual"){
      model_matrix[,which(grepl(input$discipline, colnames(model_matrix)))] = 1
    }
    model_matrix[,37] = ifelse(input$discipline == "sprint" | input$discipline == "individual", 0,1)
    
    # setup lap
    model_matrix[,17] = as.numeric(input$shooting_nr)
    model_matrix[,18] = ifelse((as.numeric(input$shooting_nr) %% 2) == 1, 0,1)
    model_matrix[,21] = model_matrix[,20] + (as.numeric(input$shooting_nr) - 1) * 5
    model_matrix[,22] = model_matrix[,21] / (ifelse(input$discipline == "sprint", 10,20))
    model_matrix[,24] = 0
    model_matrix[5,24] = ifelse(model_matrix[5,22] == 1,1,0)
    
    # compute model hit probabilities and setup outputs
    hits = rbinom(5,1,predict(model, model_matrix))
    probs = predict(model, model_matrix)
    probs = round(probs, digits = 4) * 100
    probs = paste(probs,"%",sep="")
    output_text = paste("Shot 1: ",probs[1],", Shot 2: ",probs[2],", Shot 3: ",probs[3],
                      ", Shot 4: ",probs[4],", Shot 5: ",probs[5], sep = "")
    output$probs <- renderText({output_text})
    output$preds <- renderText({predict(model, model_matrix) > threshold})
    output$hits <- renderText({hits})
    output$simulate <- renderText({"Simulated outcome:"})
    output$probs_title <- renderText({"Predicted probabilities: "})
    
    # render the image output according to the simulated outcomes
    output$shot1 = renderImage({
      if (hits[1] == 1) {
        return(list(
          src = "www/biathlon_target_hit.jpg",
          contentType = "image/jpg",
          width="100px"
        ))
      } else{
        return(list(
          src = "www/biathlon_target_miss.jpg",
          filetype = "image/jpg",
          width="100px"
        ))
      }
    }, deleteFile = FALSE)
    
    output$shot2 = renderImage({
      if (hits[2] == 1) {
        return(list(
          src = "www/biathlon_target_hit.jpg",
          contentType = "image/jpg",
          width="100px"
        ))
      } else{
        return(list(
          src = "www/biathlon_target_miss.jpg",
          filetype = "image/jpg",
          width="100px"
        ))
      }
    }, deleteFile = FALSE)
    
    output$shot3 = renderImage({
      if (hits[3] == 1) {
        return(list(
          src = "www/biathlon_target_hit.jpg",
          contentType = "image/jpg",
          width="100px"
        ))
      } else{
        return(list(
          src = "www/biathlon_target_miss.jpg",
          filetype = "image/jpg",
          width="100px"
        ))
      }
    }, deleteFile = FALSE)
    
    output$shot4 = renderImage({
      if (hits[4] == 1) {
        return(list(
          src = "www/biathlon_target_hit.jpg",
          contentType = "image/jpg",
          width="100px"
        ))
      } else{
        return(list(
          src = "www/biathlon_target_miss.jpg",
          filetype = "image/jpg",
          width="100px"
        ))
      }
    }, deleteFile = FALSE)
    
    output$shot5 = renderImage({
      if (hits[5] == 1) {
        return(list(
          src = "www/biathlon_target_hit.jpg",
          contentType = "image/jpg",
          width="100px"
        ))
      } else{
        return(list(
          src = "www/biathlon_target_miss.jpg",
          filetype = "image/jpg",
          width="100px"
        ))
      }
    }, deleteFile = FALSE)
    
  })
})
