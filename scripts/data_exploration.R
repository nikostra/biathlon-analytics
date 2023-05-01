library(openxlsx)
library(tidyxl)
library(readxl)
library(tidyverse)

read_shooting_series = function(file, gender){
  shooting_series <- read_excel(file, sheet = "Protocol")
  shooting_series$shooting_time_1 = round(as.numeric(format(as.POSIXct(shooting_series$`Shot 1`, format = "%Y-%m-%d %H:%M:%OS"), "%OS2")), digits = 1)
  shooting_series$shooting_time_1[!grepl("\\.0$", format(shooting_series$shooting_time_1))] <- shooting_series$shooting_time_1[!grepl("\\.0$", format(shooting_series$shooting_time_1))] - 1
  shooting_series$shooting_time_2 = round(as.numeric(format(as.POSIXct(shooting_series$`Shot 2`, format = "%Y-%m-%d %H:%M:%OS"), "%OS2")), digits = 1)
  shooting_series$shooting_time_2[!grepl("\\.0$", format(shooting_series$shooting_time_2))] <- shooting_series$shooting_time_2[!grepl("\\.0$", format(shooting_series$shooting_time_2))] - 1
  shooting_series$shooting_time_3 = round(as.numeric(format(as.POSIXct(shooting_series$`Shot 3`, format = "%Y-%m-%d %H:%M:%OS"), "%OS2")), digits = 1)
  shooting_series$shooting_time_3[!grepl("\\.0$", format(shooting_series$shooting_time_3))] <- shooting_series$shooting_time_3[!grepl("\\.0$", format(shooting_series$shooting_time_3))] - 1
  shooting_series$shooting_time_4 = round(as.numeric(format(as.POSIXct(shooting_series$`Shot 4`, format = "%Y-%m-%d %H:%M:%OS"), "%OS2")), digits = 1)
  shooting_series$shooting_time_4[!grepl("\\.0$", format(shooting_series$shooting_time_4))] <- shooting_series$shooting_time_4[!grepl("\\.0$", format(shooting_series$shooting_time_4))] - 1
  shooting_series$shooting_time_5 = round(as.numeric(format(as.POSIXct(shooting_series$`Shot 5`, format = "%Y-%m-%d %H:%M:%OS"), "%OS2")), digits = 1)
  shooting_series$shooting_time_5[!grepl("\\.0$", format(shooting_series$shooting_time_5))] <- shooting_series$shooting_time_5[!grepl("\\.0$", format(shooting_series$shooting_time_5))] - 1
  
  athletes <<- add_athletes(shooting_series$Name,gender)
  
  shooting_series$shot1 = 1
  shooting_series$shot2 = 1
  shooting_series$shot3 = 1
  shooting_series$shot4 = 1
  shooting_series$shot5 = 1
  shooting_series <- shooting_series %>% rowwise() %>% 
    mutate(athlete_id = get_id(Name))
  
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
  
  return(shooting_series)
}

# split data from shooting series to single shots
split_series = function(shooting_series, shots, location, season, discipline){
  total_shots = ifelse(discipline == "sprint", 10, 20)
  for (series in 1:nrow(shooting_series)) {
    i = nrow(shots) + 1
    shots[i,]$target = shooting_series[series,]$shot1
    shots[i,]$shooting_time = shooting_series[series,]$shooting_time_1
    shots[i,]$shot_number_series = 1
    shots[i,]$shot_number_race = 1 + (shooting_series[series,]$Lap - 1) * 5
    shots[i,]$shot_number_race_scaled = shots[i,]$shot_number_race / total_shots
    
    shots[i+1,]$target = shooting_series[series,]$shot2
    shots[i+1,]$shooting_time = shooting_series[series,]$shooting_time_2
    shots[i+1,]$shot_number_series = 2
    shots[i+1,]$shot_number_race = 2 + (shooting_series[series,]$Lap - 1) * 5
    shots[i+1,]$shot_number_race_scaled = shots[i+1,]$shot_number_race / total_shots
    
    shots[i+2,]$target = shooting_series[series,]$shot3
    shots[i+2,]$shooting_time = shooting_series[series,]$shooting_time_3
    shots[i+2,]$shot_number_series = 3
    shots[i+2,]$shot_number_race = 3 + (shooting_series[series,]$Lap - 1) * 5
    shots[i+2,]$shot_number_race_scaled = shots[i+2,]$shot_number_race / total_shots
    
    shots[i+3,]$target = shooting_series[series,]$shot4
    shots[i+3,]$shooting_time = shooting_series[series,]$shooting_time_4
    shots[i+3,]$shot_number_series = 4
    shots[i+3,]$shot_number_race = 4 + (shooting_series[series,]$Lap - 1) * 5
    shots[i+3,]$shot_number_race_scaled = shots[i+3,]$shot_number_race / total_shots
    
    shots[i+4,]$target = shooting_series[series,]$shot5
    shots[i+4,]$shooting_time = shooting_series[series,]$shooting_time_5
    shots[i+4,]$shot_number_series = 5
    shots[i+4,]$shot_number_race = 5 + (shooting_series[series,]$Lap - 1) * 5
    shots[i+4,]$shot_number_race_scaled = shots[i+4,]$shot_number_race / total_shots
    
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

load_files(files){
  athletes = data.frame(id=c(1), name=c("wright campbell"), gender=c("m"))
  shots = data.frame(target = numeric(), location = character(), discipline = character(), season = character(),
                     athlete_id = numeric(), gender = character(), lap = numeric(), mode = character(),
                     lane = numeric(), shot_number_series = numeric(), shot_number_race = numeric(),
                     shot_number_race_scaled = numeric(), shooting_time = numeric())
  
  for (file in files){
    sho
  }
}

file = "data/22_23/Hochfilzen/SprintMen.xlsx"
shooting_series = read_shooting_series(file,"m")
shots = split_series(shooting_series,shots,"hochfilzen", "2223", "sprint")

file = "data/22_23/Hochfilzen/Sprint Women.xlsx"
shooting_series = read_shooting_series(file,"f")
shots = split_series(shooting_series,shots,"hochfilzen", "2223", "sprint")

file = "data/22_23/Hochfilzen/Pursuit Men.xlsx"
shooting_series = read_shooting_series(file,"m")
shots = split_series(shooting_series,shots,"hochfilzen", "2223", "pursuit")

file = "data/22_23/Hochfilzen/Pursuit Women.xlsx"
shooting_series = read_shooting_series(file,"f")
shots = split_series(shooting_series,shots,"hochfilzen", "2223", "pursuit")

file_2 = "data/22_23/Antholz/Sprint Men.xlsx"
shooting_series = read_shooting_series(file_2,"m")
shots = split_series(shooting_series,shots,"antholz", "2223", "sprint")

file_2 = "data/22_23/Antholz/Pursuit Men.xlsx"
shooting_series = read_shooting_series(file_2,"m")
shots = split_series(shooting_series,shots,"antholz", "2223", "pursuit")

file_2 = "data/22_23/Antholz/Sprint Women.xlsx"
shooting_series = read_shooting_series(file_2,"f")
shots = split_series(shooting_series,shots,"antholz", "2223", "sprint")

file_3 = "data/22_23/Antholz/Pursuit Women.xlsx"
shooting_series = read_shooting_series(file_3,"f")
shots = split_series(shooting_series,shots,"antholz", "2223", "pursuit")

file_5 = "data/20_21/Oberhof 2/Mass Start Women.xlsx"
shooting_series = read_shooting_series(file_5,"f")
shots = split_series(shooting_series,shots,"oberhof", "2021", "mass start")

file_6 = "data/22_23/Nove Mesto na Morave/Sprint Women.xlsx"
shooting_series = read_shooting_series(file_6,"f")
shots = split_series(shooting_series,shots,"nove mesto", "2223", "sprint")

file_7 = "data/22_23/Nove Mesto na Morave/Pursuit Women.xlsx"
shooting_series = read_shooting_series(file_7,"f")
shots = split_series(shooting_series,shots,"nove mesto", "2223", "pursuit")

file_8 = "data/22_23/Nove Mesto na Morave/Pursuit Men.xlsx"
shooting_series = read_shooting_series(file_8,"m")
shots = split_series(shooting_series,shots,"nove mesto", "2223", "pursuit")

file_9 = "data/22_23/Nove Mesto na Morave/Sprint Men.xlsx"
shooting_series = read_shooting_series(file_9,"m")
shots = split_series(shooting_series,shots,"nove mesto", "2223", "sprint")

file_10 = "data/22_23/Pokljuka/Pursuit Women.xlsx"
shooting_series = read_shooting_series(file_10,"f")
shots = split_series(shooting_series,shots,"pokljuka", "2223", "pursuit")

file_11 = "data/22_23/Pokljuka/Pursuit Men.xlsx"
shooting_series = read_shooting_series(file_11,"m")
shots = split_series(shooting_series,shots,"pokljuka", "2223", "pursuit")

file_12 = "data/22_23/Pokljuka/SprintMen.xlsx"
shooting_series = read_shooting_series(file_12,"m")
shots = split_series(shooting_series,shots,"pokljuka", "2223", "sprint")

file_13 = "data/22_23/Pokljuka/Sprint Women.xlsx"
shooting_series = read_shooting_series(file_13,"f")
shots = split_series(shooting_series,shots,"pokljuka", "2223", "sprint")

file_13 = "data/22_23/Ruhpolding/Individual Women.xlsx"
shooting_series = read_shooting_series(file_13,"f")
shots = split_series(shooting_series,shots,"pokljuka", "2223", "individual")

file_13 = "data/22_23/Ruhpolding/Individual Men.xlsx"
shooting_series = read_shooting_series(file_13,"m")
shots = split_series(shooting_series,shots,"pokljuka", "2223", "individual")

file_13 = "data/22_23/Ruhpolding/Mass Start Women.xlsx"
shooting_series = read_shooting_series(file_13,"f")
shots = split_series(shooting_series,shots,"pokljuka", "2223", "mass start")

file_13 = "data/22_23/Ruhpolding/Mass Start Men.xlsx"
shooting_series = read_shooting_series(file_13,"m")
shots = split_series(shooting_series,shots,"pokljuka", "2223", "mass start")

file_o = "data/22_23/Oberhof WCH/Mass Start Men.xlsx"
shooting_series = read_shooting_series(file_o,"m")
shots = split_series(shooting_series,shots,"oberhof", "2223", "mass start")

file_o = "data/22_23/Oberhof WCH/Mass Start Women.xlsx"
shooting_series = read_shooting_series(file_o,"f")
shots = split_series(shooting_series,shots,"oberhof", "2223", "mass start")

file_o = "data/22_23/Oberhof WCH/Individual Men.xlsx"
shooting_series = read_shooting_series(file_o,"m")
shots = split_series(shooting_series,shots,"oberhof", "2223", "individual")

file_o = "data/22_23/Oberhof WCH/Individual Women.xlsx"
shooting_series = read_shooting_series(file_o,"f")
shots = split_series(shooting_series,shots,"oberhof", "2223", "individual")

file_o = "data/22_23/Oberhof WCH/Sprint Men.xlsx"
shooting_series = read_shooting_series(file_o,"m")
shots = split_series(shooting_series,shots,"oberhof", "2223", "sprint")

file_o = "data/22_23/Oberhof WCH/Sprint Women.xlsx"
shooting_series = read_shooting_series(file_o,"f")
shots = split_series(shooting_series,shots,"oberhof", "2223", "sprint")

file_o = "data/22_23/Oberhof WCH/Pursuit Men.xlsx"
shooting_series = read_shooting_series(file_o,"m")
shots = split_series(shooting_series,shots,"oberhof", "2223", "pursuit")

file_o = "data/22_23/Oberhof WCH/Pursuit Women.xlsx"
shooting_series = read_shooting_series(file_o,"f")
shots = split_series(shooting_series,shots,"oberhof", "2223", "pursuit")

#season 21/22



shots$location = as.factor(shots$location)
shots$discipline = as.factor(shots$discipline)
shots$season = as.factor(shots$season)
shots$gender = as.factor(shots$gender)
shots$mode = as.factor(shots$mode)

summary(shots)

# try regression
y = as.factor(shots$target)
regData = shots[,2:13]
model = glm(y ~ ., data = regData)
summary(model)
max(na.omit(predict(model, shots)))

library(xgboost)

