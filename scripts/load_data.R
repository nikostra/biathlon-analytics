library(openxlsx)
library(tidyxl)
library(readxl)

read_shooting_series = function(file, gender){
  shooting_series <- read_excel(file, sheet = "Protocol")
  
  shooting_series$shot1 = 1
  shooting_series$shot2 = 1
  shooting_series$shot3 = 1
  shooting_series$shot4 = 1
  shooting_series$shot5 = 1
  
  cells = xlsx_cells(file, sheets = 'Protocol')
  cells = cells[25:nrow(cells), ]
  formats = xlsx_formats(file)
  misses = cells[cells$local_format_id %in% which(formats$local$font$bold),] # get all cells with bold letters -> misses
  # next step: add miss to previous data frame
  for (i in 1:nrow(misses)){
    shot_number = misses[i,]$col - 7
    column_name = paste("shot",shot_number, sep="")
    shooting_series[[misses[i,]$row - 1, column_name]] = 0
  }
  
  rows_to_remove = c()
  if(any(grepl("==", shooting_series$`Shot 1`))){
    rows_to_remove = c(rows_to_remove, grep("==", shooting_series$`Shot 1`))
    shooting_series$shooting_time_1 = round(as.numeric(shooting_series$`Shot 1`)*86400, digits = 1)    
  } else {
    shooting_series$shooting_time_1 = round(as.numeric(format(as.POSIXct(shooting_series$`Shot 1`, format = "%Y-%m-%d %H:%M:%OS"), "%OS2")), digits = 1)
    shooting_series$shooting_time_1[!grepl("\\.0$", format(shooting_series$shooting_time_1))] <- shooting_series$shooting_time_1[!grepl("\\.0$", format(shooting_series$shooting_time_1))] - 1
  }
  
  if(any(grepl("==", shooting_series$`Shot 2`))){
    rows_to_remove = c(rows_to_remove, grep("==", shooting_series$`Shot 2`))
    shooting_series$shooting_time_2 = round(as.numeric(shooting_series$`Shot 2`)*86400, digits = 1)    
  } else {
    shooting_series$shooting_time_2 = round(as.numeric(format(as.POSIXct(shooting_series$`Shot 2`, format = "%Y-%m-%d %H:%M:%OS"), "%OS2")), digits = 1)
    shooting_series$shooting_time_2[!grepl("\\.0$", format(shooting_series$shooting_time_2))] <- shooting_series$shooting_time_2[!grepl("\\.0$", format(shooting_series$shooting_time_2))] - 1
  }
  
  if(any(grepl("==", shooting_series$`Shot 3`))){
    rows_to_remove = c(rows_to_remove, grep("==", shooting_series$`Shot 3`))
    shooting_series$shooting_time_3 = round(as.numeric(shooting_series$`Shot 3`)*86400, digits = 1)    
  } else {
    shooting_series$shooting_time_3 = round(as.numeric(format(as.POSIXct(shooting_series$`Shot 3`, format = "%Y-%m-%d %H:%M:%OS"), "%OS2")), digits = 1)
    shooting_series$shooting_time_3[!grepl("\\.0$", format(shooting_series$shooting_time_3))] <- shooting_series$shooting_time_3[!grepl("\\.0$", format(shooting_series$shooting_time_3))] - 1
  }
  
  if(any(grepl("==", shooting_series$`Shot 4`))){
    rows_to_remove = c(rows_to_remove, grep("==", shooting_series$`Shot 4`))
    shooting_series$shooting_time_4 = round(as.numeric(shooting_series$`Shot 4`)*86400, digits = 1)    
  } else {
    shooting_series$shooting_time_4 = round(as.numeric(format(as.POSIXct(shooting_series$`Shot 4`, format = "%Y-%m-%d %H:%M:%OS"), "%OS2")), digits = 1)
    shooting_series$shooting_time_4[!grepl("\\.0$", format(shooting_series$shooting_time_4))] <- shooting_series$shooting_time_4[!grepl("\\.0$", format(shooting_series$shooting_time_4))] - 1
  }
  
  if(any(grepl("==", shooting_series$`Shot 5`))){
    rows_to_remove = c(rows_to_remove, grep("==", shooting_series$`Shot 5`))
    shooting_series$shooting_time_5 = round(as.numeric(shooting_series$`Shot 5`)*86400, digits = 1)    
  } else {
    shooting_series$shooting_time_5 = round(as.numeric(format(as.POSIXct(shooting_series$`Shot 5`, format = "%Y-%m-%d %H:%M:%OS"), "%OS2")), digits = 1)
    shooting_series$shooting_time_5[!grepl("\\.0$", format(shooting_series$shooting_time_5))] <- shooting_series$shooting_time_5[!grepl("\\.0$", format(shooting_series$shooting_time_5))] - 1
  }
  
  if(length(rows_to_remove) > 0){
    shooting_series = shooting_series[-unique(rows_to_remove),] # remove rows with columns with errors
  }

  athletes <<- add_athletes(shooting_series$Name,gender)
  
  shooting_series <- shooting_series %>% rowwise() %>% 
    mutate(athlete_id = get_id(Name))
  
  shooting_series$race_id = race_id
  race_id <<- race_id + 1
  return(shooting_series)
}

# split data from shooting series to single shots
split_series = function(shooting_series, location, season, discipline){
  shots = get_shots_df()
  total_shots = ifelse(discipline == "sprint", 10, 20)
  for (series in 1:nrow(shooting_series)) {
    i = nrow(shots) + 1
    
    shots[i:(i+4),]$race_id = shooting_series[series,]$race_id
    
    shots[i,]$shooting_time = shooting_series[series,]$shooting_time_1
    shots[i,]$shot_number_series = 1
    shots[i,]$shot_number_race = 1 + (shooting_series[series,]$Lap - 1) * 5
    shots[i,]$shot_number_race_scaled = shots[i,]$shot_number_race / total_shots
    shots[i,]$last_shot = 0
    # calculate number of prior hits: 
    prior_hits = shots %>% filter(race_id == shooting_series[series,]$race_id & athlete_id == shooting_series[series,]$athlete_id) %>% pull(target) %>% na.omit %>% sum
    shots[i,]$hitrate_in_race = ifelse(shots[i,]$shot_number_race == 1,1, prior_hits/(shots[i,]$shot_number_race - 1))
    shots[i,]$target = shooting_series[series,]$shot1
    
    shots[i+1,]$target = shooting_series[series,]$shot2
    shots[i+1,]$shooting_time = shooting_series[series,]$shooting_time_2
    shots[i+1,]$shot_number_series = 2
    shots[i+1,]$shot_number_race = 2 + (shooting_series[series,]$Lap - 1) * 5
    shots[i+1,]$shot_number_race_scaled = shots[i+1,]$shot_number_race / total_shots
    shots[i+1,]$last_shot = 0
    prior_hits = prior_hits + (1*shots[i,]$target)
    shots[i+1,]$hitrate_in_race = prior_hits/(shots[i,]$shot_number_race)
    
    shots[i+2,]$target = shooting_series[series,]$shot3
    shots[i+2,]$shooting_time = shooting_series[series,]$shooting_time_3
    shots[i+2,]$shot_number_series = 3
    shots[i+2,]$shot_number_race = 3 + (shooting_series[series,]$Lap - 1) * 5
    shots[i+2,]$shot_number_race_scaled = shots[i+2,]$shot_number_race / total_shots
    shots[i+2,]$last_shot = 0
    prior_hits = prior_hits + (1*shots[i+1,]$target)
    shots[i+2,]$hitrate_in_race = prior_hits/(shots[i+1,]$shot_number_race)
    
    shots[i+3,]$target = shooting_series[series,]$shot4
    shots[i+3,]$shooting_time = shooting_series[series,]$shooting_time_4
    shots[i+3,]$shot_number_series = 4
    shots[i+3,]$shot_number_race = 4 + (shooting_series[series,]$Lap - 1) * 5
    shots[i+3,]$shot_number_race_scaled = shots[i+3,]$shot_number_race / total_shots
    shots[i+3,]$last_shot = 0
    prior_hits = prior_hits + (1*shots[i+2,]$target)
    shots[i+3,]$hitrate_in_race = prior_hits/(shots[i+2,]$shot_number_race)
    
    shots[i+4,]$target = shooting_series[series,]$shot5
    shots[i+4,]$shooting_time = shooting_series[series,]$shooting_time_5
    shots[i+4,]$shot_number_series = 5
    shots[i+4,]$shot_number_race = 5 + (shooting_series[series,]$Lap - 1) * 5
    shots[i+4,]$shot_number_race_scaled = shots[i+4,]$shot_number_race / total_shots
    shots[i+4,]$last_shot = ifelse(shots[i+4,]$shot_number_race_scaled == 1, 1, 0)
    prior_hits = prior_hits + (1*shots[i+3,]$target)
    shots[i+4,]$hitrate_in_race = prior_hits/(shots[i+3,]$shot_number_race)
    
    shots[i:(i+4),]$location = location
    shots[i:(i+4),]$discipline = discipline
    shots[i:(i+4),]$season = season
    shots[i:(i+4),]$athlete_id = shooting_series[series,]$athlete_id
    shots[i:(i+4),]$gender = athletes[which(athletes$id == shooting_series[series,]$athlete_id),]$gender
    shots[i:(i+4),]$lap = shooting_series[series,]$Lap
    shots[i:(i+4),]$mode = shooting_series[series,]$Mode
    shots[i:(i+4),]$lane = shooting_series[series,]$Lane
  }
  return(shots)
}


get_shots_df = function(){
  shots = data.frame(target = numeric(), location = character(), discipline = character(), season = character(),
                     athlete_id = numeric(), gender = character(), lap = numeric(), mode = character(),
                     lane = numeric(), shot_number_series = numeric(), shot_number_race = numeric(),
                     shot_number_race_scaled = numeric(), shooting_time = numeric(), last_shot = numeric(), 
                     hitrate_in_race = numeric(), race_id = numeric(), wind_speed = numeric(),
                     wind_cat = character(), snow_cond = character(), weather = character(),
                     air_temp = numeric())
  return(shots)
}

add_weather = function(shots, wind_speed, snow_cond, weather, air_temp){
  shots$wind_speed = wind_speed
  if(wind_speed < 1.75){
    shots$wind_cat = "calm"
  } else if(wind_speed < 3){
    shots$wind_cat = "moderate"
  } else {
    shots$wind_cat = "strong"
  }
  shots$snow_cond = snow_cond
  shots$weather = weather
  shots$air_temp = air_temp
  return(shots)
}

load_shots = function(load_all = TRUE){
  shots = get_shots_df()
  print("Load season 2022-2023")
  
  file = "data/22_23/Hochfilzen/SprintMen.xlsx"
  shooting_series = read_shooting_series(file,"m")
  new_shots = split_series(shooting_series,"hochfilzen", "2223", "sprint")
  new_shots = add_weather(new_shots, 0.1, "hard packed", "cloudy", -1.0)
  shots = rbind(shots, new_shots)
  
  file = "data/22_23/Hochfilzen/Sprint Women.xlsx"
  shooting_series = read_shooting_series(file,"f")
  new_shots = split_series(shooting_series,"hochfilzen", "2223", "sprint")
  new_shots = add_weather(new_shots, 1.5, "hard packed", "cloudy", -0.7)
  shots = rbind(shots, new_shots)
  
  file = "data/22_23/Hochfilzen/Pursuit Men.xlsx"
  shooting_series = read_shooting_series(file,"m")
  new_shots = split_series(shooting_series,"hochfilzen", "2223", "pursuit")
  new_shots = add_weather(new_shots, 1.2, "hard packed", "snowing", -5.9)
  shots = rbind(shots, new_shots)
  
  file = "data/22_23/Hochfilzen/Pursuit Women.xlsx"
  shooting_series = read_shooting_series(file,"f")
  new_shots = split_series(shooting_series,"hochfilzen", "2223", "pursuit")
  new_shots = add_weather(new_shots, 1.6, "hard packed", "snowing", -1.0)
  shots = rbind(shots, new_shots)
  
  file_2 = "data/22_23/Antholz/Sprint Men.xlsx"
  shooting_series = read_shooting_series(file_2,"m")
  new_shots = split_series(shooting_series,"antholz", "2223", "sprint")
  new_shots = add_weather(new_shots, 1.2, "hard packed", "cloudy", -6.2)
  shots = rbind(shots, new_shots)
  
  file_2 = "data/22_23/Antholz/Pursuit Men.xlsx"
  shooting_series = read_shooting_series(file_2,"m")
  new_shots = split_series(shooting_series,"antholz", "2223", "pursuit")
  new_shots = add_weather(new_shots, 1.3, "hard packed", "cloudy", -5.5)
  shots = rbind(shots, new_shots)
  
  file_2 = "data/22_23/Antholz/Sprint Women.xlsx"
  shooting_series = read_shooting_series(file_2,"f")
  new_shots = split_series(shooting_series,"antholz", "2223", "sprint")
  new_shots = add_weather(new_shots, 0.9, "hard packed", "cloudy", -6.0)
  shots = rbind(shots, new_shots)
  
  file_3 = "data/22_23/Antholz/Pursuit Women.xlsx"
  shooting_series = read_shooting_series(file_3,"f")
  new_shots = split_series(shooting_series,"antholz", "2223", "pursuit")
  new_shots = add_weather(new_shots, 1.3, "hard packed", "cloudy", -5.5)
  shots = rbind(shots, new_shots)
  
  file_6 = "data/22_23/Nove Mesto na Morave/Sprint Women.xlsx"
  shooting_series = read_shooting_series(file_6,"f")
  new_shots = split_series(shooting_series,"nove mesto", "2223", "sprint")
  new_shots = add_weather(new_shots, 1.5, "hard packed", "sunny", 4.5)
  shots = rbind(shots, new_shots)
  
  file_7 = "data/22_23/Nove Mesto na Morave/Pursuit Women.xlsx"
  shooting_series = read_shooting_series(file_7,"f")
  new_shots = split_series(shooting_series,"nove mesto", "2223", "pursuit")
  new_shots = add_weather(new_shots, 2.1, "hard packed", "cloudy", 0.1)
  shots = rbind(shots, new_shots)
  
  file_8 = "data/22_23/Nove Mesto na Morave/Pursuit Men.xlsx"
  shooting_series = read_shooting_series(file_8,"m")
  new_shots = split_series(shooting_series,"nove mesto", "2223", "pursuit")
  new_shots = add_weather(new_shots, 2.1, "hard packed", "cloudy", 0.0)
  shots = rbind(shots, new_shots)
  
  file_9 = "data/22_23/Nove Mesto na Morave/Sprint Men.xlsx"
  shooting_series = read_shooting_series(file_9,"m")
  new_shots = split_series(shooting_series,"nove mesto", "2223", "sprint")
  new_shots = add_weather(new_shots, 0.6, "hard packed", "sunny", 6.4)
  shots = rbind(shots, new_shots)
  
  file_10 = "data/22_23/Pokljuka/Pursuit Women.xlsx"
  shooting_series = read_shooting_series(file_10,"f")
  new_shots = split_series(shooting_series,"pokljuka", "2223", "pursuit")
  new_shots = add_weather(new_shots, 1.0, "hard packed", "sunny", 2.9)
  shots = rbind(shots, new_shots)
  
  file_11 = "data/22_23/Pokljuka/Pursuit Men.xlsx"
  shooting_series = read_shooting_series(file_11,"m")
  new_shots = split_series(shooting_series,"pokljuka", "2223", "pursuit")
  new_shots = add_weather(new_shots, 0.8, "hard packed", "sunny", 0.8)
  shots = rbind(shots, new_shots)
  
  file_12 = "data/22_23/Pokljuka/SprintMen.xlsx"
  shooting_series = read_shooting_series(file_12,"m")
  new_shots = split_series(shooting_series,"pokljuka", "2223", "sprint")
  new_shots = add_weather(new_shots, 0.4, "hard packed", "sunny", 2.7)
  shots = rbind(shots, new_shots)
  
  file_13 = "data/22_23/Pokljuka/Sprint Women.xlsx"
  shooting_series = read_shooting_series(file_13,"f")
  new_shots = split_series(shooting_series,"pokljuka", "2223", "sprint")
  new_shots = add_weather(new_shots, 0.2, "wet", "cloudy", 2.5)
  shots = rbind(shots, new_shots)
  
  file_13 = "data/22_23/Ruhpolding/Individual Women.xlsx"
  shooting_series = read_shooting_series(file_13,"f")
  new_shots = split_series(shooting_series,"pokljuka", "2223", "individual")
  new_shots = add_weather(new_shots, 0.4, "hard packed", "cloudy", 3.3)
  shots = rbind(shots, new_shots)
  
  file_13 = "data/22_23/Ruhpolding/Individual Men.xlsx"
  shooting_series = read_shooting_series(file_13,"m")
  new_shots = split_series(shooting_series,"pokljuka", "2223", "individual")
  new_shots = add_weather(new_shots, 1.6, "hard packed", "cloudy", 2.7)
  shots = rbind(shots, new_shots)
  
  file_13 = "data/22_23/Ruhpolding/Mass Start Women.xlsx"
  shooting_series = read_shooting_series(file_13,"f")
  new_shots = split_series(shooting_series,"pokljuka", "2223", "mass start")
  new_shots = add_weather(new_shots, 2.0, "hard packed", "rainy", 3.4)
  shots = rbind(shots, new_shots)
  
  file_13 = "data/22_23/Ruhpolding/Mass Start Men.xlsx"
  shooting_series = read_shooting_series(file_13,"m")
  new_shots = split_series(shooting_series,"pokljuka", "2223", "mass start")
  new_shots = add_weather(new_shots, 2.0, "hard packed", "cloudy", 6.9)
  shots = rbind(shots, new_shots)
  
  print("Load WCH 2023")
  
  file_o = "data/22_23/Oberhof WCH/Mass Start Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"oberhof", "2223", "mass start")
  new_shots = add_weather(new_shots, 1.2, "hard packed", "rainy", 3.7)
  shots = rbind(shots, new_shots)
  
  file_o = "data/22_23/Oberhof WCH/Mass Start Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"oberhof", "2223", "mass start")
  new_shots = add_weather(new_shots, 1.0, "hard packed", "rainy", 3.5)
  shots = rbind(shots, new_shots)
  
  file_o = "data/22_23/Oberhof WCH/Individual Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"oberhof", "2223", "individual")
  new_shots = add_weather(new_shots, 0.6, "hard packed", "sunny", 8.9)
  shots = rbind(shots, new_shots)
  
  file_o = "data/22_23/Oberhof WCH/Individual Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"oberhof", "2223", "individual")
  new_shots = add_weather(new_shots, 2.6, "hard packed", "sunny", 6.5)
  shots = rbind(shots, new_shots)
  
  file_o = "data/22_23/Oberhof WCH/Sprint Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"oberhof", "2223", "sprint")
  new_shots = add_weather(new_shots, 1.4, "hard packed", "cloudy", 0.2)
  shots = rbind(shots, new_shots)
  
  file_o = "data/22_23/Oberhof WCH/Sprint Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"oberhof", "2223", "sprint")
  new_shots = add_weather(new_shots, 0.9, "hard packed", "cloudy", -1.2)
  shots = rbind(shots, new_shots)
  
  file_o = "data/22_23/Oberhof WCH/Pursuit Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"oberhof", "2223", "pursuit")
  new_shots = add_weather(new_shots, 1.6, "hard packed", "cloudy", 4.1)
  shots = rbind(shots, new_shots)
  
  file_o = "data/22_23/Oberhof WCH/Pursuit Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"oberhof", "2223", "pursuit")
  new_shots = add_weather(new_shots, 0.9, "hard packed", "cloudy", 4.3)
  shots = rbind(shots, new_shots)
  
  if(load_all){
  
  #season 21/22
  print("Load season 2021-22")
  file_o = "data/21_22/Antholz/Individual Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"antholz", "2122", "individual")
  new_shots = add_weather(new_shots, 1.3, "hard packed", "sunny", -0.8)
  shots = rbind(shots, new_shots)
  
  file_o = "data/21_22/Antholz/Individual Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"antholz", "2122", "individual")
  new_shots = add_weather(new_shots, 0.9, "hard packed", "sunny", -3.3)
  shots = rbind(shots, new_shots)
  
  file_o = "data/21_22/Antholz/Mass Start Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"antholz", "2122", "mass start")
  new_shots = add_weather(new_shots, 1.5, "hard packed", "snowing", 0.5)
  shots = rbind(shots, new_shots)
  
  file_o = "data/21_22/Antholz/Mass Start Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"antholz", "2122", "mass start")
  new_shots = add_weather(new_shots, 0.8, "hard packed", "sunny", 0.2)
  shots = rbind(shots, new_shots)
  
  
  file_o = "data/21_22/Hochfilzen/Sprint Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"hochfilzen", "2122", "sprint")
  new_shots = add_weather(new_shots, 4.2, "hard packed", "cloudy", -4.5)
  shots = rbind(shots, new_shots)
  
  file_o = "data/21_22/Hochfilzen/Sprint Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"hochfilzen", "2122", "sprint")
  new_shots = add_weather(new_shots, 4.2, "hard packed", "cloudy", -3.9)
  shots = rbind(shots, new_shots)
  
  file_o = "data/21_22/Hochfilzen/Pursuit Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"hochfilzen", "2122", "pursuit")
  new_shots = add_weather(new_shots, 1.2, "hard packed", "snowing", -1.6)
  shots = rbind(shots, new_shots)
  
  file_o = "data/21_22/Hochfilzen/Pursuit Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"hochfilzen", "2122", "pursuit")
  new_shots = add_weather(new_shots, 1.0, "hard packed", "cloudy", -1.2)
  shots = rbind(shots, new_shots)
  
  
  
  file_o = "data/21_22/Oberhof/Sprint Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"oberhof", "2122", "sprint")
  new_shots = add_weather(new_shots, 2.2, "hard packed", "snowing", -3.3)
  shots = rbind(shots, new_shots)
  
  file_o = "data/21_22/Oberhof/Sprint Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"oberhof", "2122", "sprint")
  new_shots = add_weather(new_shots, 1.0, "hard packed", "fog", -3.6)
  shots = rbind(shots, new_shots)
  
  file_o = "data/21_22/Oberhof/Pursuit Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"oberhof", "2122", "pursuit")
  new_shots = add_weather(new_shots, 1.6, "hard packed", "snowing", -0.9)
  shots = rbind(shots, new_shots)
  
  file_o = "data/21_22/Oberhof/Pursuit Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"oberhof", "2122", "pursuit")
  new_shots = add_weather(new_shots, 1.1, "hard packed", "snowing", -1.0)
  shots = rbind(shots, new_shots)
  
  
  file_o = "data/21_22/Ruhpolding/Sprint Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"ruhpolding", "2122", "sprint")
  new_shots = add_weather(new_shots, 0.3, "hard packed", "sunny", -3.7)
  shots = rbind(shots, new_shots)
  
  file_o = "data/21_22/Ruhpolding/Sprint Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"ruhpolding", "2122", "sprint")
  new_shots = add_weather(new_shots, 0.2, "hard packed", "sunny", -7.0)
  shots = rbind(shots, new_shots)
  
  file_o = "data/21_22/Ruhpolding/Pursuit Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"ruhpolding", "2122", "pursuit")
  new_shots = add_weather(new_shots, 2.1, "hard packed", "sunny", 2.5)
  shots = rbind(shots, new_shots)
  
  file_o = "data/21_22/Ruhpolding/Pursuit Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"ruhpolding", "2122", "pursuit")
  new_shots = add_weather(new_shots, 1.2, "hard packed", "sunny", 0.8)
  shots = rbind(shots, new_shots)
  
  # season 2020-21
  print("Load season 2020-21")
  file_o = "data/20_21/Antholz/Individual Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"antholz", "2021", "individual")
  new_shots = add_weather(new_shots, 0.3, "hard packed", "cloudy", 1.3)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Antholz/Individual Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"antholz", "2021", "individual")
  new_shots = add_weather(new_shots, 0.7, "hard packed", "snowing", -1.1)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Antholz/Mass Start Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"antholz", "2021", "mass start")
  new_shots = add_weather(new_shots, 0.3, "hard packed", "sunny", -5.2)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Antholz/Mass Start Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"antholz", "2021", "mass start")
  new_shots = add_weather(new_shots, 1.1, "hard packed", "snowing", -2.6)
  shots = rbind(shots, new_shots)
  
  
  file_o = "data/20_21/Hochfilzen 1/Sprint Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"hochfilzen", "2021", "sprint")
  new_shots = add_weather(new_shots, 2.2, "hard packed", "cloudy", -1.3)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Hochfilzen 1/Sprint Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"hochfilzen", "2021", "sprint")
  new_shots = add_weather(new_shots, 1.8, "hard packed", "cloudy", -2.1)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Hochfilzen 1/Pursuit Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"hochfilzen", "2021", "pursuit")
  new_shots = add_weather(new_shots, 0.4, "hard packed", "snowing", 0.1)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Hochfilzen 1/Pursuit Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"hochfilzen", "2021", "pursuit")
  new_shots = add_weather(new_shots, 0.0, "hard packed", "snowing", 2.4)
  shots = rbind(shots, new_shots)
  
  
  file_o = "data/20_21/Hochfilzen 2/Sprint Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"hochfilzen", "2021", "sprint")
  new_shots = add_weather(new_shots, 0.9, "hard packed", "sunny", 5.3)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Hochfilzen 2/Sprint Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"hochfilzen", "2021", "sprint")
  new_shots = add_weather(new_shots, 0.8, "hard packed", "sunny", 3.8)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Hochfilzen 2/Pursuit Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"hochfilzen", "2021", "pursuit")
  new_shots = add_weather(new_shots, 0.9, "hard packed", "sunny", 3.3)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Hochfilzen 2/Pursuit Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"hochfilzen", "2021", "pursuit")
  new_shots = add_weather(new_shots, 0.3, "hard packed", "sunny", 3.2)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Hochfilzen 2/Mass Start Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"hochfilzen", "2021", "mass start")
  new_shots = add_weather(new_shots, 1.0, "hard packed", "sunny", 2.9)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Hochfilzen 2/Mass Start Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"hochfilzen", "2021", "mass start")
  new_shots = add_weather(new_shots, 0.3, "hard packed", "cloudy", 1.7)
  shots = rbind(shots, new_shots)
  
  print("Load NMNM 2021")
  file_6 = "data/20_21/Nove Mesto na Morave 1/Sprint Women.xlsx"
  shooting_series = read_shooting_series(file_6,"f")
  new_shots = split_series(shooting_series,"nove mesto", "2021", "sprint")
  new_shots = add_weather(new_shots, 3.0, "hard packed", "cloudy", 0.9)
  shots = rbind(shots, new_shots)
  
  file_7 = "data/20_21/Nove Mesto na Morave 1/Pursuit Women.xlsx"
  shooting_series = read_shooting_series(file_7,"f")
  new_shots = split_series(shooting_series,"nove mesto", "2021", "pursuit")
  new_shots = add_weather(new_shots, 2.0, "hard packed", "sunny", 3.1)
  shots = rbind(shots, new_shots)
  
  file_8 = "data/20_21/Nove Mesto na Morave 1/Pursuit Men.xlsx"
  shooting_series = read_shooting_series(file_8,"m")
  new_shots = split_series(shooting_series,"nove mesto", "2021", "pursuit")
  new_shots = add_weather(new_shots, 1.2, "hard packed", "sunny", 3.8)
  shots = rbind(shots, new_shots)
  
  file_9 = "data/20_21/Nove Mesto na Morave 1/Sprint Men.xlsx"
  shooting_series = read_shooting_series(file_9,"m")
  new_shots = split_series(shooting_series,"nove mesto", "2021", "sprint")
  new_shots = add_weather(new_shots, 1.6, "hard packed", "cloudy", 1.3)
  shots = rbind(shots, new_shots)
  
  
  file_6 = "data/20_21/Nove Mesto na Morave 2/Sprint Women.xlsx"
  shooting_series = read_shooting_series(file_6,"f")
  new_shots = split_series(shooting_series,"nove mesto", "2021", "sprint")
  new_shots = add_weather(new_shots, 0.6, "hard packed", "rainy", 0.4)
  shots = rbind(shots, new_shots)
  
  file_7 = "data/20_21/Nove Mesto na Morave 2/Pursuit Women.xlsx"
  shooting_series = read_shooting_series(file_7,"f")
  new_shots = split_series(shooting_series,"nove mesto", "2021", "pursuit")
  new_shots = add_weather(new_shots, 1.2, "hard packed", "cloudy", 7.2)
  shots = rbind(shots, new_shots)
  
  file_8 = "data/20_21/Nove Mesto na Morave 2/Pursuit Men.xlsx"
  shooting_series = read_shooting_series(file_8,"m")
  new_shots = split_series(shooting_series,"nove mesto", "2021", "pursuit")
  new_shots = add_weather(new_shots, 1.6, "hard packed", "cloudy", 8.7)
  shots = rbind(shots, new_shots)
  
  file_9 = "data/20_21/Nove Mesto na Morave 2/Sprint Men.xlsx"
  shooting_series = read_shooting_series(file_9,"m")
  new_shots = split_series(shooting_series,"nove mesto", "2021", "sprint")
  new_shots = add_weather(new_shots, 1.0, "hard packed", "raining", 0.3)
  shots = rbind(shots, new_shots)
  
  
  file_o = "data/20_21/Oberhof 1/Sprint Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"oberhof", "2021", "sprint")
  new_shots = add_weather(new_shots, 1.1, "powder", "snowing", -2.8)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Oberhof 1/Sprint Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"oberhof", "2021", "sprint")
  new_shots = add_weather(new_shots, 0.7, "powder", "snowing", -2.9)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Oberhof 1/Pursuit Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"oberhof", "2021", "pursuit")
  new_shots = add_weather(new_shots, 0.4, "powder", "cloudy", -3.8)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Oberhof 1/Pursuit Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"oberhof", "2021", "pursuit")
  new_shots = add_weather(new_shots, 0.6, "powder", "cloudy", -3.7)
  shots = rbind(shots, new_shots)
  
  
  file_o = "data/20_21/Oberhof 2/Sprint Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"oberhof", "2021", "sprint")
  new_shots = add_weather(new_shots, 1.9, "powder", "cloudy", -3.1)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Oberhof 2/Sprint Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"oberhof", "2021", "sprint")
  new_shots = add_weather(new_shots, 1.1, "powder", "cloudy", -3.4)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Oberhof 2/Mass Start Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"oberhof", "2021", "mass start")
  new_shots = add_weather(new_shots, 0.7, "powder", "cloudy", -5.6)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Oberhof 2/Mass Start Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"oberhof", "2021", "mass start")
  new_shots = add_weather(new_shots, 1.9, "powder", "snowing", -5.6)
  shots = rbind(shots, new_shots)
  
  print("Load 2021 WCH")
  file_o = "data/20_21/Pokljuka WCH/Individual Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"pokljuka", "2021", "individual")
  new_shots = add_weather(new_shots, 2.3, "hard packed", "sunny", 2.6)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Pokljuka WCH/Individual Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"pokljuka", "2021", "individual")
  new_shots = add_weather(new_shots, 1.1, "hard packed", "sunny", 3.4)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Pokljuka WCH/Sprint Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"pokljuka", "2021", "sprint")
  new_shots = add_weather(new_shots, 2.9, "hard packed", "sunny", -10.8)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Pokljuka WCH/Sprint Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"pokljuka", "2021", "sprint")
  new_shots = add_weather(new_shots, 0.4, "hard packed", "cloudy", -10.9)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Pokljuka WCH/Pusuit Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"pokljuka", "2021", "pursuit")
  new_shots = add_weather(new_shots, 1.3, "hard packed", "sunny", -7.7)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Pokljuka WCH/Pusuit Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"pokljuka", "2021", "pursuit")
  new_shots = add_weather(new_shots, 1.9, "hard packed", "sunny", -5.3)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Pokljuka WCH/Mass Start Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"pokljuka", "2021", "mass start")
  new_shots = add_weather(new_shots, 0.7, "hard packed", "sunny", 5.0)
  shots = rbind(shots, new_shots)
  
  file_o = "data/20_21/Pokljuka WCH/Mass Start Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"pokljuka", "2021", "mass start")
  new_shots = add_weather(new_shots, 1.0, "hard packed", "sunny", 3.9)
  shots = rbind(shots, new_shots)
  }
  
  shots$location = as.factor(shots$location)
  shots$discipline = as.factor(shots$discipline)
  shots$season = as.factor(shots$season)
  shots$gender = as.factor(shots$gender)
  shots$mode = as.factor(shots$mode)
  shots$weather = as.factor(shots$weather)
  shots$snow_cond = as.factor(shots$snow_cond)
  shots$wind_cat = as.factor(shots$wind_cat)
  
  return(shots)
}
