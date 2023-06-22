
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

  athletes <<- add_athletes(shooting_series$Name,gender,shooting_series$Nation)
  
  shooting_series <- shooting_series %>% rowwise() %>% 
    mutate(athlete_id = get_id(Name))
  
  shooting_series$race_id = race_id
  race_id <<- race_id + 1
  return(shooting_series)
}

# split data from shooting series to single shots
split_series = function(shooting_series, location, season, discipline, level=1){
  shots = get_shots_df()
  total_shots = ifelse(discipline == "sprint", 10, 20)
  for (series in 1:nrow(shooting_series)) {
    i = nrow(shots) + 1
    athlete_db_id = which(athletes$id == shooting_series[series,]$athlete_id)
    
    shots[i:(i+4),]$race_id = shooting_series[series,]$race_id
    
    shots[i,]$shooting_time = shooting_series[series,]$shooting_time_1
    shots[i,]$shot_number_series = 1
    shots[i,]$shot_number_race = 1 + (shooting_series[series,]$Lap - 1) * 5
    shots[i,]$shot_number_race_scaled = shots[i,]$shot_number_race / total_shots
    shots[i,]$last_shot = 0
    shots[i,]$long_shooting_time = 0
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
    shots[i+1,]$long_shooting_time = ifelse(shots[i+1,]$shooting_time > 3.8,1,0) # hit rate decreases a lot with longer shooting time
    prior_hits = prior_hits + (1*shots[i,]$target)
    shots[i+1,]$hitrate_in_race = prior_hits/(shots[i,]$shot_number_race)
    
    shots[i+2,]$target = shooting_series[series,]$shot3
    shots[i+2,]$shooting_time = shooting_series[series,]$shooting_time_3
    shots[i+2,]$shot_number_series = 3
    shots[i+2,]$shot_number_race = 3 + (shooting_series[series,]$Lap - 1) * 5
    shots[i+2,]$shot_number_race_scaled = shots[i+2,]$shot_number_race / total_shots
    shots[i+2,]$last_shot = 0
    shots[i+2,]$long_shooting_time = ifelse(shots[i+2,]$shooting_time > 3.8,1,0) # hit rate decreases a lot with longer shooting time
    prior_hits = prior_hits + (1*shots[i+1,]$target)
    shots[i+2,]$hitrate_in_race = prior_hits/(shots[i+1,]$shot_number_race)
    
    shots[i+3,]$target = shooting_series[series,]$shot4
    shots[i+3,]$shooting_time = shooting_series[series,]$shooting_time_4
    shots[i+3,]$shot_number_series = 4
    shots[i+3,]$shot_number_race = 4 + (shooting_series[series,]$Lap - 1) * 5
    shots[i+3,]$shot_number_race_scaled = shots[i+3,]$shot_number_race / total_shots
    shots[i+3,]$last_shot = 0
    shots[i+3,]$long_shooting_time = ifelse(shots[i+3,]$shooting_time > 3.8,1,0) # hit rate decreases a lot with longer shooting time
    prior_hits = prior_hits + (1*shots[i+2,]$target)
    shots[i+3,]$hitrate_in_race = prior_hits/(shots[i+2,]$shot_number_race)
    
    shots[i+4,]$target = shooting_series[series,]$shot5
    shots[i+4,]$shooting_time = shooting_series[series,]$shooting_time_5
    shots[i+4,]$shot_number_series = 5
    shots[i+4,]$shot_number_race = 5 + (shooting_series[series,]$Lap - 1) * 5
    shots[i+4,]$shot_number_race_scaled = shots[i+4,]$shot_number_race / total_shots
    shots[i+4,]$last_shot = ifelse(shots[i+4,]$shot_number_race_scaled == 1, 1, 0)
    shots[i+4,]$long_shooting_time = ifelse(shots[i+4,]$shooting_time > 3.8,1,0) # hit rate decreases a lot with longer shooting time
    prior_hits = prior_hits + (1*shots[i+3,]$target)
    shots[i+4,]$hitrate_in_race = prior_hits/(shots[i+3,]$shot_number_race)
    
    shots[i:(i+4),]$location = location
    shots[i:(i+4),]$discipline = discipline
    shots[i:(i+4),]$head_to_head = ifelse(discipline == "sprint" | discipline == "individual",0,1)
    shots[i:(i+4),]$season = season
    shots[i:(i+4),]$athlete_id = shooting_series[series,]$athlete_id
    shots[i:(i+4),]$gender = athletes[athlete_db_id,]$gender
    shots[i:(i+4),]$lap = shooting_series[series,]$Lap
    shots[i:(i+4),]$mode = shooting_series[series,]$Mode
    shots[i:(i+4),]$lane = shooting_series[series,]$Lane
    shots[i:(i+4),]$competition_level = level # 1 for WC, 0 for IBU Cup
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
                     air_temp = numeric(), head_to_head = numeric(), long_shooting_time = numeric(),
                     competition_level = numeric())
  return(shots)
}

add_weather = function(shots, wind_speed, snow_cond, weather, air_temp){
  shots$wind_speed = wind_speed
  if(wind_speed < 2){
    shots$wind_cat = "calm"
  } else if(wind_speed < 3.5){
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
  
  if(load_all){
    
    #season 2018-2019
    print("Load season 2018-19")
    
    file_o = "data/18_19/Pokljuka/Individual Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"pokljuka", "1819", "individual")
    new_shots = add_weather(new_shots, 0.9, "compact", "cloudy", 0.7)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Pokljuka/Individual Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"pokljuka", "1819", "individual")
    new_shots = add_weather(new_shots, 0.7, "compact", "cloudy", 2.5)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Pokljuka/Sprint Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"oberhof", "1819", "sprint")
    new_shots = add_weather(new_shots, 0.2, "compact", "cloudy", 2.2)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Pokljuka/Sprint Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"oberhof", "1819", "sprint")
    new_shots = add_weather(new_shots, 0.1, "compact", "cloudy", -1.5)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Pokljuka/Pursuit Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"oberhof", "1819", "pursuit")
    new_shots = add_weather(new_shots, 0.8, "compact", "snowing", 1.6)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Pokljuka/Pursuit Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"oberhof", "1819", "pursuit")
    new_shots = add_weather(new_shots, 0.3, "compact", "cloudy", -0.1)
    shots = rbind(shots, new_shots)
    
    
    file_o = "data/18_19/Hochfilzen/Sprint Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"hochfilzen", "1819", "sprint")
    new_shots = add_weather(new_shots, 0.5, "hard packed", "sunny", -6.5)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Hochfilzen/Sprint Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"hochfilzen", "1819", "sprint")
    new_shots = add_weather(new_shots, 0.3, "hard packed", "sunny", -2.9)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Hochfilzen/Pursuit Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"hochfilzen", "1819", "pursuit")
    new_shots = add_weather(new_shots, 0.5, "hard packed", "sunny", -5.9)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Hochfilzen/Pursuit Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"hochfilzen", "1819", "pursuit")
    new_shots = add_weather(new_shots, 1.6, "hard packed", "cloudy", -6.7)
    shots = rbind(shots, new_shots)
    
    
    
    file_o = "data/18_19/Nove Mesto/Sprint Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"nove mesto", "1819", "sprint")
    new_shots = add_weather(new_shots, 0.7, "hard packed", "cloudy", -3.7)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Nove Mesto/Sprint Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"nove mesto", "1819", "sprint")
    new_shots = add_weather(new_shots, 0.6, "hard packed", "snowing", -1.2)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Nove Mesto/Pursuit Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"nove mesto", "1819", "pursuit")
    new_shots = add_weather(new_shots, 2.3, "hard packed", "cloudy", 5.7)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Nove Mesto/Pursuit Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"nove mesto", "1819", "pursuit")
    new_shots = add_weather(new_shots, 3.9, "hard packed", "raining", 4.0)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Nove Mesto/Mass Start Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"nove mesto", "1819", "mass start")
    new_shots = add_weather(new_shots, 1.9, "hard packed", "cloudy", 3.8)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Nove Mesto/Mass Start Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"nove mesto", "1819", "mass start")
    new_shots = add_weather(new_shots, 0.8, "hard packed", "cloudy", 3.9)
    shots = rbind(shots, new_shots)
    
    
    print("Load 2019 Oberhof")
    
    file_o = "data/18_19/Oberhof/Sprint Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"oberhof", "1819", "sprint")
    new_shots = add_weather(new_shots, 1.4, "hard packed", "snowing", -3.5)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Oberhof/Sprint Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"oberhof", "1819", "sprint")
    new_shots = add_weather(new_shots, 0.6, "hard packed", "cloudy", -5.6)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Oberhof/Pursuit Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"oberhof", "1819", "pursuit")
    new_shots = add_weather(new_shots, 1.9, "hard packed", "snowing", -0.4)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Oberhof/Pursuit Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"oberhof", "1819", "pursuit")
    new_shots = add_weather(new_shots, 3.5, "hard packed", "snowing", -0.2)
    shots = rbind(shots, new_shots)
    
    
    file_o = "data/18_19/Ruhpolding/Sprint Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"ruhpolding", "1819", "sprint")
    new_shots = add_weather(new_shots, 1.6, "compact", "sunny", 0.1)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Ruhpolding/Sprint Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"ruhpolding", "1819", "sprint")
    new_shots = add_weather(new_shots, 1.0, "compact", "sunny", 0.9)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Ruhpolding/Mass Start Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"ruhpolding", "1819", "mass start")
    new_shots = add_weather(new_shots, 0.9, "compact", "cloudy", -3.8)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Ruhpolding/Mass Start Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"ruhpolding", "1819", "mass start")
    new_shots = add_weather(new_shots, 0.7, "compact", "cloudy", -1.2)
    shots = rbind(shots, new_shots)
    
    
    file_o = "data/18_19/Antholz/Sprint Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"antholz", "1819", "sprint")
    new_shots = add_weather(new_shots, 1.1, "compact", "sunny", -4.4)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Antholz/Sprint Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"antholz", "1819", "sprint")
    new_shots = add_weather(new_shots, 1.2, "compact", "sunny", -4.6)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Antholz/Pursuit Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"antholz", "1819", "pursuit")
    new_shots = add_weather(new_shots, 1.1, "compact", "cloudy", 2.8)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Antholz/Pursuit Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"antholz", "1819", "pursuit")
    new_shots = add_weather(new_shots, 4.7, "compact", "raining", 4.2)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Antholz/Mass Start Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"antholz", "1819", "mass start")
    new_shots = add_weather(new_shots, 0.2, "compact", "cloudy", -2.2)
    shots = rbind(shots, new_shots)
    
    file_o = "data/18_19/Antholz/Mass Start Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"antholz", "1819", "mass start")
    new_shots = add_weather(new_shots, 0.7, "compact", "cloudy", -0.1)
    shots = rbind(shots, new_shots)
    

    # season 2019-20
    print("Load season 2019-20")
    
    file_o = "data/19_20/Hochfilzen/Sprint Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"hochfilzen", "1920", "sprint")
    new_shots = add_weather(new_shots, 0.5, "hard packed", "cloudy", -3.5)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Hochfilzen/Sprint Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"hochfilzen", "1920", "sprint")
    new_shots = add_weather(new_shots, 0.5, "hard packed", "snowing", -1.3)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Hochfilzen/Pursuit Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"hochfilzen", "1920", "pursuit")
    new_shots = add_weather(new_shots, 0.8, "hard packed", "cloudy", 1.0)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Hochfilzen/Pursuit Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"hochfilzen", "1920", "pursuit")
    new_shots = add_weather(new_shots, 0.4, "hard packed", "cloudy", 0.1)
    shots = rbind(shots, new_shots)
    
    
    file_o = "data/19_20/Oberhof/Sprint Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"oberhof", "1920", "sprint")
    new_shots = add_weather(new_shots, 2.4, "hard packed", "cloudy", 3.8)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Oberhof/Sprint Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"oberhof", "1920", "sprint")
    new_shots = add_weather(new_shots, 1.8, "hard packed", "raining", 5.4)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Oberhof/Mass Start Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"oberhof", "1920", "mass start")
    new_shots = add_weather(new_shots, 2.6, "hard packed", "cloudy", -2.2)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Oberhof/Mass Start Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"oberhof", "1920", "mass start")
    new_shots = add_weather(new_shots, 3.5, "hard packed", "cloudy", -2.5)
    shots = rbind(shots, new_shots)
    
    
    file_10 = "data/19_20/Osrblie IBU Cup/Individual Women.xlsx"
    shooting_series = read_shooting_series(file_10,"f")
    new_shots = split_series(shooting_series,"brezno", "1920", "individual", 0)
    new_shots = add_weather(new_shots, 0.9, "hard packed", "cloudy", -1.2)
    shots = rbind(shots, new_shots)
    
    file_11 = "data/19_20/Osrblie IBU Cup/Individual Men.xlsx"
    shooting_series = read_shooting_series(file_11,"m")
    new_shots = split_series(shooting_series,"brezno", "1920", "individual", 0)
    new_shots = add_weather(new_shots, 0.3, "hard packed", "cloudy", 0.4)
    shots = rbind(shots, new_shots)
    
    file_10 = "data/19_20/Osrblie IBU Cup/Sprint Women.xlsx"
    shooting_series = read_shooting_series(file_10,"f")
    new_shots = split_series(shooting_series,"brezno", "1920", "sprint", 0)
    new_shots = add_weather(new_shots, 0.6, "hard packed", "sunny", -2.2)
    shots = rbind(shots, new_shots)
    
    file_11 = "data/19_20/Osrblie IBU Cup/Sprint Men.xlsx"
    shooting_series = read_shooting_series(file_11,"m")
    new_shots = split_series(shooting_series,"brezno", "1920", "sprint", 0)
    new_shots = add_weather(new_shots, 0.7, "hard packed", "sunny", -1.5)
    shots = rbind(shots, new_shots)
    
    
    file_10 = "data/19_20/Osrblie IBU Cup 2/Sprint women.xlsx"
    shooting_series = read_shooting_series(file_10,"f")
    new_shots = split_series(shooting_series,"brezno", "1920", "sprint", 0)
    new_shots = add_weather(new_shots, 0.3, "hard packed", "sunny", 3.1)
    shots = rbind(shots, new_shots)
    
    file_11 = "data/19_20/Osrblie IBU Cup 2/Sprint men.xlsx"
    shooting_series = read_shooting_series(file_11,"m")
    new_shots = split_series(shooting_series,"brezno", "1920", "sprint", 0)
    new_shots = add_weather(new_shots, 0.2, "hard packed", "sunny", 1.7)
    shots = rbind(shots, new_shots)
    
    
    
    file_o = "data/19_20/Ruhpolding/Sprint Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"ruhpolding", "1920", "sprint")
    new_shots = add_weather(new_shots, 0.7, "compact", "sunny", 3.4)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Ruhpolding/Sprint Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"ruhpolding", "1920", "sprint")
    new_shots = add_weather(new_shots, 0.9, "hard packed", "sunny", 2.6)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Ruhpolding/Pursuit Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"ruhpolding", "1920", "pursuit")
    new_shots = add_weather(new_shots, 0.5, "hard packed", "cloudy", -0.5)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Ruhpolding/Pursuit Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"ruhpolding", "1920", "pursuit")
    new_shots = add_weather(new_shots, 0.4, "hard packed", "snowing", -1.2)
    shots = rbind(shots, new_shots)
    
    
    
    file_o = "data/19_20/Pokljuka/Individual men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"pokljuka", "1920", "individual")
    new_shots = add_weather(new_shots, 0.4, "hard packed", "snowing", -1.2)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Pokljuka/Individual women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"pokljuka", "1920", "individual")
    new_shots = add_weather(new_shots, 0.7, "hard packed", "sunny", 2.7)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Pokljuka/Mass Start men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"oberhof", "1920", "mass start")
    new_shots = add_weather(new_shots, 0.9, "hard packed", "cloudy", 3.3)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Pokljuka/Mass Start women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"oberhof", "1920", "mass start")
    new_shots = add_weather(new_shots, 0.7, "hard packed", "cloudy", 1.7)
    shots = rbind(shots, new_shots)
    
    
    print("load 2020 WCH")
    
    file_o = "data/19_20/Antholz WCH/Sprint Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"antholz", "1920", "sprint")
    new_shots = add_weather(new_shots, 2.9, "hard packed", "cloudy", 3.4)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Antholz WCH/Sprint Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"antholz", "1920", "sprint")
    new_shots = add_weather(new_shots, 0.6, "compact", "cloudy", 4.9)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Antholz WCH/Pursuit Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"antholz", "1920", "pursuit")
    new_shots = add_weather(new_shots, 0.6, "compact", "cloudy", 6.9)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Antholz WCH/Pursuit Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"antholz", "1920", "pursuit")
    new_shots = add_weather(new_shots, 1.0, "compact", "cloudy", 4.8)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Antholz WCH/Individual Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"antholz", "1920", "individual")
    new_shots = add_weather(new_shots, 1.5, "compact", "sunny", 4.9)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Antholz WCH/Individual Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"antholz", "1920", "individual")
    new_shots = add_weather(new_shots, 3.5, "compact", "snowing", 1.4)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Antholz WCH/Mass Start Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"antholz", "1920", "mass start")
    new_shots = add_weather(new_shots, 1.5, "compact", "cloudy", 8.9)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Antholz WCH/Mass Start Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"antholz", "1920", "mass start")
    new_shots = add_weather(new_shots, 1.3, "compact", "raining", 5.6)
    shots = rbind(shots, new_shots)
    
    
    file_o = "data/19_20/Nove Mesto/Sprint Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"nove mesto", "1920", "sprint")
    new_shots = add_weather(new_shots, 0.5, "hard packed", "cloudy", 2.4)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Nove Mesto/Sprint Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"nove mesto", "1920", "sprint")
    new_shots = add_weather(new_shots, 1.9, "hard packed", "cloudy", 4.0)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Nove Mesto/Mass Start Men.xlsx"
    shooting_series = read_shooting_series(file_o,"m")
    new_shots = split_series(shooting_series,"nove mesto", "1920", "mass start")
    new_shots = add_weather(new_shots, 1.9, "hard packed", "cloudy", 4.8)
    shots = rbind(shots, new_shots)
    
    file_o = "data/19_20/Nove Mesto/Mass Start Women.xlsx"
    shooting_series = read_shooting_series(file_o,"f")
    new_shots = split_series(shooting_series,"nove mesto", "1920", "mass start")
    new_shots = add_weather(new_shots, 1.6, "hard packed", "cloudy", 4.8)
    shots = rbind(shots, new_shots)
    

    
    # season 2020-21
    print("Load season 2020-21")
    
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
    
    file_9 = "data/20_21/Nove Mesto na Morave 1/Sprint Men.xlsx"
    shooting_series = read_shooting_series(file_9,"m")
    new_shots = split_series(shooting_series,"nove mesto", "2021", "sprint")
    new_shots = add_weather(new_shots, 1.6, "hard packed", "cloudy", 1.3)
    shots = rbind(shots, new_shots)
    
    file_8 = "data/20_21/Nove Mesto na Morave 1/Pursuit Men.xlsx"
    shooting_series = read_shooting_series(file_8,"m")
    new_shots = split_series(shooting_series,"nove mesto", "2021", "pursuit")
    new_shots = add_weather(new_shots, 1.2, "hard packed", "sunny", 3.8)
    shots = rbind(shots, new_shots)
    
    
    file_6 = "data/20_21/Nove Mesto na Morave 2/Sprint Women.xlsx"
    shooting_series = read_shooting_series(file_6,"f")
    new_shots = split_series(shooting_series,"nove mesto", "2021", "sprint")
    new_shots = add_weather(new_shots, 0.6, "hard packed", "raining", 0.4)
    shots = rbind(shots, new_shots)
    
    file_7 = "data/20_21/Nove Mesto na Morave 2/Pursuit Women.xlsx"
    shooting_series = read_shooting_series(file_7,"f")
    new_shots = split_series(shooting_series,"nove mesto", "2021", "pursuit")
    new_shots = add_weather(new_shots, 1.2, "hard packed", "cloudy", 7.2)
    shots = rbind(shots, new_shots)
    
    file_9 = "data/20_21/Nove Mesto na Morave 2/Sprint Men.xlsx"
    shooting_series = read_shooting_series(file_9,"m")
    new_shots = split_series(shooting_series,"nove mesto", "2021", "sprint")
    new_shots = add_weather(new_shots, 1.0, "hard packed", "raining", 0.3)
    shots = rbind(shots, new_shots)
    
    file_8 = "data/20_21/Nove Mesto na Morave 2/Pursuit Men.xlsx"
    shooting_series = read_shooting_series(file_8,"m")
    new_shots = split_series(shooting_series,"nove mesto", "2021", "pursuit")
    new_shots = add_weather(new_shots, 1.6, "hard packed", "cloudy", 8.7)
    shots = rbind(shots, new_shots)
    

    print("Load 2021 WCH")

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
    
    
    
    file_10 = "data/20_21/Osrblie IBU Cup 1/Osrblie IBUCup 13.2.2021 women sprint.xlsx"
    shooting_series = read_shooting_series(file_10,"f")
    new_shots = split_series(shooting_series,"brezno", "2021", "sprint", 0)
    new_shots = add_weather(new_shots, 1.4, "hard packed", "cloudy", -4.6)
    shots = rbind(shots, new_shots)
    
    file_11 = "data/20_21/Osrblie IBU Cup 1/Osrblie IBUCup 13.2.2021 men sprint.xlsx"
    shooting_series = read_shooting_series(file_11,"m")
    new_shots = split_series(shooting_series,"brezno", "2021", "sprint", 0)
    new_shots = add_weather(new_shots, 3.5, "hard packed", "snowing", 1.4)
    shots = rbind(shots, new_shots)
    
    file_10 = "data/20_21/Osrblie IBU Cup 1/Osrblie IBUCup 14.2.2021 women pursuit.xlsx"
    shooting_series = read_shooting_series(file_10,"f")
    new_shots = split_series(shooting_series,"brezno", "2021", "pursuit", 0)
    new_shots = add_weather(new_shots, 1.1, "hard packed", "cloudy", 0.2)
    shots = rbind(shots, new_shots)
    
    file_11 = "data/20_21/Osrblie IBU Cup 1/Osrblie IBUCup 14.2.2021 men pursuit.xlsx"
    shooting_series = read_shooting_series(file_11,"m")
    new_shots = split_series(shooting_series,"brezno", "2021", "pursuit", 0)
    new_shots = add_weather(new_shots, 1.1, "hard packed", "cloudy", -2.7)
    shots = rbind(shots, new_shots)
    
    
    
    file_10 = "data/20_21/Osrblie IBU Cup 2/Osrblie IBUCup 17.2.2021 short individual men.xlsx"
    shooting_series = read_shooting_series(file_10,"f")
    new_shots = split_series(shooting_series,"brezno", "2021", "individual", 0)
    new_shots = add_weather(new_shots, 0.3, "hard packed", "snowing", -2.6)
    shots = rbind(shots, new_shots)
    
    file_11 = "data/20_21/Osrblie IBU Cup 2/Osrblie IBUCup 17.2.2021 short individual women.xlsx"
    shooting_series = read_shooting_series(file_11,"m")
    new_shots = split_series(shooting_series,"brezno", "2021", "individual", 0)
    new_shots = add_weather(new_shots, 0.3, "hard packed", "snowing", -1.1)
    shots = rbind(shots, new_shots)
    
    file_10 = "data/20_21/Osrblie IBU Cup 2/Osrblie IBUCup 20.2.2021 sprint women.xlsx"
    shooting_series = read_shooting_series(file_10,"f")
    new_shots = split_series(shooting_series,"brezno", "2021", "sprint", 0)
    new_shots = add_weather(new_shots, 0.9, "hard packed", "cloudy", 2.6)
    shots = rbind(shots, new_shots)
    
    file_11 = "data/20_21/Osrblie IBU Cup 2/Osrblie IBUCup 20.2.2021 sprint men.xlsx"
    shooting_series = read_shooting_series(file_11,"m")
    new_shots = split_series(shooting_series,"brezno", "2021", "sprint", 0)
    new_shots = add_weather(new_shots, 0.9, "hard packed", "cloudy", 4.2)
    shots = rbind(shots, new_shots)
    
    file_10 = "data/20_21/Osrblie IBU Cup 2/Osrblie IBUCup 21.2.2021 women pursuit.xlsx"
    shooting_series = read_shooting_series(file_10,"f")
    new_shots = split_series(shooting_series,"brezno", "2021", "pursuit", 0)
    new_shots = add_weather(new_shots, 0.2, "hard packed", "cloudy", 0.8)
    shots = rbind(shots, new_shots)
    
    file_11 = "data/20_21/Osrblie IBU Cup 2/Osrblie IBUCup 21.2.2021 men pursuit.xlsx"
    shooting_series = read_shooting_series(file_11,"m")
    new_shots = split_series(shooting_series,"brezno", "2021", "pursuit", 0)
    new_shots = add_weather(new_shots, 0.2, "hard packed", "cloudy", 5.4)
    shots = rbind(shots, new_shots)
    
    #season 21/22
    
    print("Load season 2021-22")
    
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
    
    
    
    file_10 = "data/21_22/Osrblie IBU CUP 1/Sprint Women.xlsx"
    shooting_series = read_shooting_series(file_10,"f")
    new_shots = split_series(shooting_series,"brezno", "2122", "sprint", 0)
    new_shots = add_weather(new_shots, 1.5, "hard packed", "cloudy", -6.6)
    shots = rbind(shots, new_shots)
    
    file_11 = "data/21_22/Osrblie IBU CUP 1/Sprint Men.xlsx"
    shooting_series = read_shooting_series(file_11,"m")
    new_shots = split_series(shooting_series,"brezno", "2122", "sprint", 0)
    new_shots = add_weather(new_shots, 0.8, "hard packed", "cloudy", -6.2)
    shots = rbind(shots, new_shots)
    
    file_10 = "data/21_22/Osrblie IBU CUP 1/Sprint Women 9.1.2022.xlsx"
    shooting_series = read_shooting_series(file_10,"f")
    new_shots = split_series(shooting_series,"brezno", "2122", "sprint", 0)
    new_shots = add_weather(new_shots, 1.4, "hard packed", "cloudy", -0.7)
    shots = rbind(shots, new_shots)
    
    file_11 = "data/21_22/Osrblie IBU CUP 1/Sprint Men 9.1.2022.xlsx"
    shooting_series = read_shooting_series(file_11,"m")
    new_shots = split_series(shooting_series,"brezno", "2122", "sprint", 0)
    new_shots = add_weather(new_shots, 0.7, "hard packed", "cloudy", -2.3)
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
    
    
    file_10 = "data/21_22/Osrblie IBU CUP 2/IBU Cup Brezno-Osrblie Short Individual Women 12.1.2022.xlsx"
    shooting_series = read_shooting_series(file_10,"f")
    new_shots = split_series(shooting_series,"brezno", "2122", "individual", 0)
    new_shots = add_weather(new_shots, 1.4, "hard packed", "cloudy", -10.2)
    shots = rbind(shots, new_shots)
    
    file_11 = "data/21_22/Osrblie IBU CUP 2/IBU Cup Brezno-Osrblie SprintMen 15.1.2022.xlsx"
    shooting_series = read_shooting_series(file_11,"m")
    new_shots = split_series(shooting_series,"brezno", "2122", "sprint", 0)
    new_shots = add_weather(new_shots, 2.9, "hard packed", "cloudy", 3.5)
    shots = rbind(shots, new_shots)
    
    file_10 = "data/21_22/Osrblie IBU CUP 2/IBU Cup Brezno-Osrblie Sprint Women 14.1.2022.xlsx"
    shooting_series = read_shooting_series(file_10,"f")
    new_shots = split_series(shooting_series,"brezno", "2122", "sprint", 0)
    new_shots = add_weather(new_shots, 1.6, "hard packed", "cloudy", -2.6)
    shots = rbind(shots, new_shots)
    
    file_11 = "data/21_22/Osrblie IBU CUP 2/IBU Cup Brezno-Osrblie Pursuit Women 15.1.2022.xlsx"
    shooting_series = read_shooting_series(file_11,"m")
    new_shots = split_series(shooting_series,"brezno", "2122", "pursuit", 0)
    new_shots = add_weather(new_shots, 2.8, "hard packed", "cloudy", 2.5)
    shots = rbind(shots, new_shots)
    
    
    
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
    
    
    
    file_10 = "data/21_22/Arber OECH/Individual Women.xlsx"
    shooting_series = read_shooting_series(file_10,"f")
    new_shots = split_series(shooting_series,"arber", "2122", "individual", 0)
    new_shots = add_weather(new_shots, 0.5, "hard packed", "fog", -3.5)
    shots = rbind(shots, new_shots)
    
    file_11 = "data/21_22/Arber OECH/Individual Men.xlsx"
    shooting_series = read_shooting_series(file_11,"m")
    new_shots = split_series(shooting_series,"arber", "2122", "individual", 0)
    new_shots = add_weather(new_shots, 0.8, "hard packed", "fog", -4.7)
    shots = rbind(shots, new_shots)
    
    file_10 = "data/21_22/Arber OECH/Sprint Women.xlsx"
    shooting_series = read_shooting_series(file_10,"f")
    new_shots = split_series(shooting_series,"arber", "2122", "sprint", 0)
    new_shots = add_weather(new_shots, 2.1, "hard packed", "snowing", -2.2)
    shots = rbind(shots, new_shots)
    
    file_11 = "data/21_22/Arber OECH/Sprint Men.xlsx"
    shooting_series = read_shooting_series(file_11,"m")
    new_shots = split_series(shooting_series,"arber", "2122", "sprint", 0)
    new_shots = add_weather(new_shots, 0.2, "hard packed", "snowing", -2.0)
    shots = rbind(shots, new_shots)
    
    file_10 = "data/21_22/Arber OECH/Pursuit Women.xlsx"
    shooting_series = read_shooting_series(file_10,"f")
    new_shots = split_series(shooting_series,"arber", "2122", "pursuit", 0)
    new_shots = add_weather(new_shots, 2.6, "hard packed", "snowing", -1.8)
    shots = rbind(shots, new_shots)
    
    file_11 = "data/21_22/Arber OECH/Pursuit Men.xlsx"
    shooting_series = read_shooting_series(file_11,"m")
    new_shots = split_series(shooting_series,"arber", "2122", "pursuit", 0)
    new_shots = add_weather(new_shots, 4.5, "hard packed", "snowing", -2.4)
    shots = rbind(shots, new_shots)
    
    
    
    file_10 = "data/21_22/Nove Mesto NM IBU CUP/Sprint Women 1.xlsx"
    shooting_series = read_shooting_series(file_10,"f")
    new_shots = split_series(shooting_series,"nove mesto", "2122", "sprint", 0)
    new_shots = add_weather(new_shots, 1.5, "hard packed", "sunny", 0.9)
    shots = rbind(shots, new_shots)
    
    file_11 = "data/21_22/Nove Mesto NM IBU CUP/Sprint Men 1.xlsx"
    shooting_series = read_shooting_series(file_11,"m")
    new_shots = split_series(shooting_series,"nove mesto", "2122", "sprint", 0)
    new_shots = add_weather(new_shots, 0.7, "hard packed", "cloudy", 2.8)
    shots = rbind(shots, new_shots)
    
    file_10 = "data/21_22/Nove Mesto NM IBU CUP/Sprint Women 2.xlsx"
    shooting_series = read_shooting_series(file_10,"f")
    new_shots = split_series(shooting_series,"nove mesto", "2122", "sprint", 0)
    new_shots = add_weather(new_shots, 2.7, "hard packed", "cloudy", 0.7)
    shots = rbind(shots, new_shots)
    
    file_11 = "data/21_22/Nove Mesto NM IBU CUP/Sprint Men 2.xlsx"
    shooting_series = read_shooting_series(file_11,"m")
    new_shots = split_series(shooting_series,"nove mesto", "2122", "sprint", 0)
    new_shots = add_weather(new_shots, 2.4, "hard packed", "cloudy", 1.1)
    shots = rbind(shots, new_shots)
    
    
    
    file_10 = "data/21_22/Ridnaun IBU Cup Finale/Sprint Women.xlsx"
    shooting_series = read_shooting_series(file_10,"f")
    new_shots = split_series(shooting_series,"ridnaun", "2122", "sprint", 0)
    new_shots = add_weather(new_shots, 0.2, "hard packed", "sunny", 9.5)
    shots = rbind(shots, new_shots)
    
    file_11 = "data/21_22/Ridnaun IBU Cup Finale/Sprint Men.xlsx"
    shooting_series = read_shooting_series(file_11,"m")
    new_shots = split_series(shooting_series,"ridnaun", "2122", "sprint", 0)
    new_shots = add_weather(new_shots, 2.1, "hard packed", "sunny", 8.4)
    shots = rbind(shots, new_shots)
    
    file_10 = "data/21_22/Ridnaun IBU Cup Finale/Pursuit Women.xlsx"
    shooting_series = read_shooting_series(file_10,"f")
    new_shots = split_series(shooting_series,"ridnaun", "2122", "pursuit", 0)
    new_shots = add_weather(new_shots, 0.1, "hard packed", "cloudy", 1.5)
    shots = rbind(shots, new_shots)
    
    file_11 = "data/21_22/Ridnaun IBU Cup Finale/Pursuit Men.xlsx"
    shooting_series = read_shooting_series(file_11,"m")
    new_shots = split_series(shooting_series,"ridnaun", "2122", "pursuit", 0)
    new_shots = add_weather(new_shots, 1.2, "hard packed", "cloudy", 3.0)
    shots = rbind(shots, new_shots)
  }
  
  
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
  
  
  
  file_12 = "data/22_23/Pokljuka IBU Cup/Short Individual Men.xlsx"
  shooting_series = read_shooting_series(file_12,"m")
  new_shots = split_series(shooting_series,"pokljuka", "2223", "individual", 0)
  new_shots = add_weather(new_shots, 0.6, "hard packed", "cloudy", -1.2)
  shots = rbind(shots, new_shots)
  
  file_13 = "data/22_23/Pokljuka IBU Cup/Short Individual Women.xlsx"
  shooting_series = read_shooting_series(file_13,"f")
  new_shots = split_series(shooting_series,"pokljuka", "2223", "individual", 0)
  new_shots = add_weather(new_shots, 0.3, "hard packed", "cloudy", 1.3)
  shots = rbind(shots, new_shots)
  
  file_10 = "data/22_23/Pokljuka IBU Cup/Sprint 1 Women.xlsx"
  shooting_series = read_shooting_series(file_10,"f")
  new_shots = split_series(shooting_series,"pokljuka", "2223", "sprint", 0)
  new_shots = add_weather(new_shots, 1.7, "hard packed", "cloudy", 2.0)
  shots = rbind(shots, new_shots)
  
  file_11 = "data/22_23/Pokljuka IBU Cup/Sprint 1 Men.xlsx"
  shooting_series = read_shooting_series(file_11,"m")
  new_shots = split_series(shooting_series,"pokljuka", "2223", "sprint", 0)
  new_shots = add_weather(new_shots, 0.8, "hard packed", "sunny", 1.5)
  shots = rbind(shots, new_shots)
  
  file_10 = "data/22_23/Pokljuka IBU Cup/Sprint 2 Women.xlsx"
  shooting_series = read_shooting_series(file_10,"f")
  new_shots = split_series(shooting_series,"pokljuka", "2223", "sprint", 0)
  new_shots = add_weather(new_shots, 2.0, "hard packed", "cloudy", 2.9)
  shots = rbind(shots, new_shots)
  
  file_11 = "data/22_23/Pokljuka IBU Cup/Sprint 2 Men.xlsx"
  shooting_series = read_shooting_series(file_11,"m")
  new_shots = split_series(shooting_series,"pokljuka", "2223", "sprint", 0)
  new_shots = add_weather(new_shots, 1.5, "hard packed", "cloudy", 1.5)
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
  new_shots = add_weather(new_shots, 2.0, "hard packed", "raining", 3.4)
  shots = rbind(shots, new_shots)
  
  file_13 = "data/22_23/Ruhpolding/Mass Start Men.xlsx"
  shooting_series = read_shooting_series(file_13,"m")
  new_shots = split_series(shooting_series,"pokljuka", "2223", "mass start")
  new_shots = add_weather(new_shots, 2.0, "hard packed", "cloudy", 6.9)
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
  
  
  print("Load WCH 2023")
  
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
  
  file_o = "data/22_23/Oberhof WCH/Mass Start Men.xlsx"
  shooting_series = read_shooting_series(file_o,"m")
  new_shots = split_series(shooting_series,"oberhof", "2223", "mass start")
  new_shots = add_weather(new_shots, 1.2, "hard packed", "raining", 3.7)
  shots = rbind(shots, new_shots)
  
  file_o = "data/22_23/Oberhof WCH/Mass Start Women.xlsx"
  shooting_series = read_shooting_series(file_o,"f")
  new_shots = split_series(shooting_series,"oberhof", "2223", "mass start")
  new_shots = add_weather(new_shots, 1.0, "hard packed", "raining", 3.5)
  shots = rbind(shots, new_shots)
  
  
  file_9 = "data/22_23/Nove Mesto na Morave/Sprint Men.xlsx"
  shooting_series = read_shooting_series(file_9,"m")
  new_shots = split_series(shooting_series,"nove mesto", "2223", "sprint")
  new_shots = add_weather(new_shots, 0.6, "hard packed", "sunny", 6.4)
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


  shots$location = as.factor(shots$location)
  shots$discipline = as.factor(shots$discipline)
  shots$season = as.factor(shots$season)
  shots$gender = as.factor(shots$gender)
  shots$mode = as.factor(shots$mode)
  shots$weather = as.factor(shots$weather)
  shots$snow_cond = as.factor(shots$snow_cond)
  shots$wind_cat = as.factor(shots$wind_cat)
  shots$competition_level = as.factor(shots$competition_level)
  
  print("compute hit rates")
  an = function(n, len) c(seq.int(to = n), rep(n, len-n))

  shots$pre_hit_rate_10 = NA
  shots$pre_hit_rate_10_mode = NA
  shots$pre_hit_rate_50 = NA
  shots$pre_hit_rate_50_mode = NA
  shots$pre_hit_rate_200 = NA
  shots$pre_hit_rate_200_mode = NA
  shots$pre_hit_rate_200_mode_shotNr = NA
  
  for (id in unique(shots$athlete_id)) {
    athlete_ids = which(shots$athlete_id == id)
    athlete_ids_p = which(shots$athlete_id == id & shots$mode == "P")
    athlete_ids_s = which(shots$athlete_id == id & shots$mode == "S")
    
    athlete_ids_s_shot1 = which(shots$athlete_id == id & shots$mode == "S" & shots$shot_number_series == 1)
    athlete_ids_s_shot2 = which(shots$athlete_id == id & shots$mode == "S" & shots$shot_number_series == 2)
    athlete_ids_s_shot3 = which(shots$athlete_id == id & shots$mode == "S" & shots$shot_number_series == 3)
    athlete_ids_s_shot4 = which(shots$athlete_id == id & shots$mode == "S" & shots$shot_number_series == 4)
    athlete_ids_s_shot5 = which(shots$athlete_id == id & shots$mode == "S" & shots$shot_number_series == 5)
    
    athlete_ids_p_shot1 = which(shots$athlete_id == id & shots$mode == "P" & shots$shot_number_series == 1)
    athlete_ids_p_shot2 = which(shots$athlete_id == id & shots$mode == "P" & shots$shot_number_series == 2)
    athlete_ids_p_shot3 = which(shots$athlete_id == id & shots$mode == "P" & shots$shot_number_series == 3)
    athlete_ids_p_shot4 = which(shots$athlete_id == id & shots$mode == "P" & shots$shot_number_series == 4)
    athlete_ids_p_shot5 = which(shots$athlete_id == id & shots$mode == "P" & shots$shot_number_series == 5)
    
    temp_shots = shots[athlete_ids,]
    temp_shots_s = shots[athlete_ids_s,]
    temp_shots_p = shots[athlete_ids_p,]
    
    temp_shots_s_shot1 = shots[athlete_ids_s_shot1,]
    temp_shots_s_shot2 = shots[athlete_ids_s_shot2,]
    temp_shots_s_shot3 = shots[athlete_ids_s_shot3,]
    temp_shots_s_shot4 = shots[athlete_ids_s_shot4,]
    temp_shots_s_shot5 = shots[athlete_ids_s_shot5,]
    
    temp_shots_p_shot1 = shots[athlete_ids_p_shot1,]
    temp_shots_p_shot2 = shots[athlete_ids_p_shot2,]
    temp_shots_p_shot3 = shots[athlete_ids_p_shot3,]
    temp_shots_p_shot4 = shots[athlete_ids_p_shot4,]
    temp_shots_p_shot5 = shots[athlete_ids_p_shot5,]
    
    #setup sequences for rolling means
    n10 = an(min(10, nrow(temp_shots)), nrow(temp_shots))
    n10_p = an(min(10, nrow(temp_shots_p)), nrow(temp_shots_p))
    n10_s = an(min(10, nrow(temp_shots_s)), nrow(temp_shots_s))
    
    n50 = an(min(50, nrow(temp_shots)), nrow(temp_shots))
    n50_p = an(min(50, nrow(temp_shots_p)), nrow(temp_shots_p))
    n50_s = an(min(50, nrow(temp_shots_s)), nrow(temp_shots_s))
    
    n200 = an(min(200, nrow(temp_shots)), nrow(temp_shots))
    n200_p = an(min(200, nrow(temp_shots_p)), nrow(temp_shots_p))
    n200_s = an(min(200, nrow(temp_shots_s)), nrow(temp_shots_s))
    
    n200_p_shot = an(min(200, nrow(temp_shots_p_shot1)), nrow(temp_shots_p_shot1))
    n200_s_shot = an(min(200, nrow(temp_shots_s_shot1)), nrow(temp_shots_s_shot1))

    # compute rolling mean and move results "up one" so that current shot is not factored in
    pre_hit_rate_10 = c(NA, frollmean(temp_shots$target, n10, adaptive=TRUE)) 
    # remove last shot
    shots$pre_hit_rate_10[athlete_ids] = pre_hit_rate_10[-length(pre_hit_rate_10)]
    
    pre_hit_rate_10_p = c(NA, frollmean(temp_shots_p$target, n10_p, adaptive=TRUE)) 
    shots$pre_hit_rate_10_mode[athlete_ids_p] = pre_hit_rate_10_p[-length(pre_hit_rate_10_p)]
    
    pre_hit_rate_10_s = c(NA, frollmean(temp_shots_s$target, n10_s, adaptive=TRUE)) 
    shots$pre_hit_rate_10_mode[athlete_ids_s] = pre_hit_rate_10_s[-length(pre_hit_rate_10_s)]
    
    
    pre_hit_rate_50 = c(NA, frollmean(temp_shots$target, n50, adaptive=TRUE)) 
    shots$pre_hit_rate_50[athlete_ids] = pre_hit_rate_50[-length(pre_hit_rate_50)]
    
    pre_hit_rate_50_p = c(NA, frollmean(temp_shots_p$target, n50_p, adaptive=TRUE)) 
    shots$pre_hit_rate_50_mode[athlete_ids_p] = pre_hit_rate_50_p[-length(pre_hit_rate_50_p)]
    
    pre_hit_rate_50_s = c(NA, frollmean(temp_shots_s$target, n50_s, adaptive=TRUE)) 
    shots$pre_hit_rate_50_mode[athlete_ids_s] = pre_hit_rate_50_s[-length(pre_hit_rate_50_s)]

    
    pre_hit_rate_200 = c(NA, frollmean(temp_shots$target, n200, adaptive=TRUE)) 
    shots$pre_hit_rate_200[athlete_ids] = pre_hit_rate_200[-length(pre_hit_rate_200)]
    
    pre_hit_rate_200_p = c(NA, frollmean(temp_shots_p$target, n200_p, adaptive=TRUE)) 
    shots$pre_hit_rate_200_mode[athlete_ids_p] = pre_hit_rate_200_p[-length(pre_hit_rate_200_p)]
    
    pre_hit_rate_200_s = c(NA, frollmean(temp_shots_s$target, n200_s, adaptive=TRUE)) 
    shots$pre_hit_rate_200_mode[athlete_ids_s] = pre_hit_rate_200_s[-length(pre_hit_rate_200_s)]
    
    pre_hit_rate_200_s_shot1 = c(NA, frollmean(temp_shots_s_shot1$target, n200_s_shot, adaptive=TRUE)) 
    shots$pre_hit_rate_200_mode_shotNr[athlete_ids_s_shot1] = pre_hit_rate_200_s_shot1[-length(pre_hit_rate_200_s_shot1)]
    pre_hit_rate_200_s_shot2 = c(NA, frollmean(temp_shots_s_shot2$target, n200_s_shot, adaptive=TRUE)) 
    shots$pre_hit_rate_200_mode_shotNr[athlete_ids_s_shot2] = pre_hit_rate_200_s_shot2[-length(pre_hit_rate_200_s_shot2)]
    pre_hit_rate_200_s_shot3 = c(NA, frollmean(temp_shots_s_shot3$target, n200_s_shot, adaptive=TRUE)) 
    shots$pre_hit_rate_200_mode_shotNr[athlete_ids_s_shot3] = pre_hit_rate_200_s_shot3[-length(pre_hit_rate_200_s_shot3)]
    pre_hit_rate_200_s_shot4 = c(NA, frollmean(temp_shots_s_shot4$target, n200_s_shot, adaptive=TRUE)) 
    shots$pre_hit_rate_200_mode_shotNr[athlete_ids_s_shot4] = pre_hit_rate_200_s_shot4[-length(pre_hit_rate_200_s_shot4)]
    pre_hit_rate_200_s_shot5 = c(NA, frollmean(temp_shots_s_shot5$target, n200_s_shot, adaptive=TRUE)) 
    shots$pre_hit_rate_200_mode_shotNr[athlete_ids_s_shot5] = pre_hit_rate_200_s_shot5[-length(pre_hit_rate_200_s_shot5)]
    
    pre_hit_rate_200_p_shot1 = c(NA, frollmean(temp_shots_p_shot1$target, n200_p_shot, adaptive=TRUE)) 
    shots$pre_hit_rate_200_mode_shotNr[athlete_ids_p_shot1] = pre_hit_rate_200_p_shot1[-length(pre_hit_rate_200_p_shot1)]
    pre_hit_rate_200_p_shot2 = c(NA, frollmean(temp_shots_p_shot2$target, n200_p_shot, adaptive=TRUE)) 
    shots$pre_hit_rate_200_mode_shotNr[athlete_ids_p_shot2] = pre_hit_rate_200_p_shot2[-length(pre_hit_rate_200_p_shot2)]
    pre_hit_rate_200_p_shot3 = c(NA, frollmean(temp_shots_p_shot3$target, n200_p_shot, adaptive=TRUE)) 
    shots$pre_hit_rate_200_mode_shotNr[athlete_ids_p_shot3] = pre_hit_rate_200_p_shot3[-length(pre_hit_rate_200_p_shot3)]
    pre_hit_rate_200_p_shot4 = c(NA, frollmean(temp_shots_p_shot4$target, n200_p_shot, adaptive=TRUE)) 
    shots$pre_hit_rate_200_mode_shotNr[athlete_ids_p_shot4] = pre_hit_rate_200_p_shot4[-length(pre_hit_rate_200_p_shot4)]
    pre_hit_rate_200_p_shot5 = c(NA, frollmean(temp_shots_p_shot5$target, n200_p_shot, adaptive=TRUE)) 
    shots$pre_hit_rate_200_mode_shotNr[athlete_ids_p_shot5] = pre_hit_rate_200_p_shot5[-length(pre_hit_rate_200_p_shot5)]
  }
  
  return(shots)
}
