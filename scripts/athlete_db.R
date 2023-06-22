athletes = data.frame(id=c(1), name=c("wright campbell"), gender=c("m"), cum_sum = c(0), shots = c(0))

# names is array with all names, gender is string "m" or "f"
add_athletes = function(names, gender, nation){
  index = 1
  for (name in names){
    name = tolower(name)
    if(!(name %in% athletes$name)){
      athletes[nrow(athletes) + 1,]$id = as.numeric(max(athletes$id)) + 1
      athletes[nrow(athletes),]$name= name
      athletes[nrow(athletes),]$gender = gender
      athletes[nrow(athletes),]$nation = nation[index]
    }
    index = index + 1
  }
  if(sum(duplicated(athletes$id)>0)){
    print("WARNING: duplicates!")
  }
  return(athletes)
}

get_id = function(name){
  id = which(athletes$name == tolower(name))
  if (length(id) == 0){
    print("WARNING: no matching name in db!")
  }
  return(id)
}


# setup athletes for shiny web app
shots_athletes = shots %>% drop_na()
athletes_data = data.frame(id=numeric(), name=character(), gender=character(), nation = character(), pre_hit_rate_10 = numeric(),             
            pre_hit_rate_10_mode_p = numeric(),pre_hit_rate_10_mode_s = numeric(), pre_hit_rate_50 = numeric(), 
            pre_hit_rate_50_mode_p = numeric(),pre_hit_rate_50_mode_s = numeric(), pre_hit_rate_200 = numeric(), 
            pre_hit_rate_200_mode_p = numeric(),pre_hit_rate_200_mode_s = numeric(), 
            pre_hit_rate_200_mode_p_shotNr_1 = numeric(),
            pre_hit_rate_200_mode_p_shotNr_2 = numeric(),
            pre_hit_rate_200_mode_p_shotNr_3 = numeric(),
            pre_hit_rate_200_mode_p_shotNr_4 = numeric(),
            pre_hit_rate_200_mode_p_shotNr_5 = numeric(),
            pre_hit_rate_200_mode_s_shotNr_1 = numeric(),
            pre_hit_rate_200_mode_s_shotNr_2 = numeric(),
            pre_hit_rate_200_mode_s_shotNr_3 = numeric(),
            pre_hit_rate_200_mode_s_shotNr_4 = numeric(),
            pre_hit_rate_200_mode_s_shotNr_5 = numeric(),
            shooting_time_p_shot1 = numeric(),
            shooting_time_p = numeric(),
            shooting_time_s_shot1 = numeric(),
            shooting_time_s = numeric()
)

for (i in 1:nrow(athletes)) {
  last_shot_of_athlete_s_1 = max(which(shots_athletes$athlete_id == athletes$id[i] & shots_athletes$mode == "S" & shots_athletes$shot_number_series == 1))
  last_shot_of_athlete_s_2 = max(which(shots_athletes$athlete_id == athletes$id[i] & shots_athletes$mode == "S" & shots_athletes$shot_number_series == 2))
  last_shot_of_athlete_s_3 = max(which(shots_athletes$athlete_id == athletes$id[i] & shots_athletes$mode == "S" & shots_athletes$shot_number_series == 3))
  last_shot_of_athlete_s_4 = max(which(shots_athletes$athlete_id == athletes$id[i] & shots_athletes$mode == "S" & shots_athletes$shot_number_series == 4))
  last_shot_of_athlete_s = max(which(shots_athletes$athlete_id == athletes$id[i] & shots_athletes$mode == "S"))
  last_shot_of_athlete_p_1 = max(which(shots_athletes$athlete_id == athletes$id[i] & shots_athletes$mode == "P" & shots_athletes$shot_number_series == 1))
  last_shot_of_athlete_p_2 = max(which(shots_athletes$athlete_id == athletes$id[i] & shots_athletes$mode == "P" & shots_athletes$shot_number_series == 2))
  last_shot_of_athlete_p_3 = max(which(shots_athletes$athlete_id == athletes$id[i] & shots_athletes$mode == "P" & shots_athletes$shot_number_series == 3))
  last_shot_of_athlete_p_4 = max(which(shots_athletes$athlete_id == athletes$id[i] & shots_athletes$mode == "P" & shots_athletes$shot_number_series == 4))
  last_shot_of_athlete_p = max(which(shots_athletes$athlete_id == athletes$id[i] & shots_athletes$mode == "P"))
  
  last_shot_of_athlete = max(last_shot_of_athlete_p, last_shot_of_athlete_s)
  
  shooting_time_p_shot1 = shots_athletes %>% filter(athlete_id == athletes$id[i] & mode == "P" & shot_number_series == 1) %>% pull(shooting_time) %>% mean()
  shooting_time_p = shots_athletes %>% filter(athlete_id == athletes$id[i] & mode == "P" & shot_number_series != 1) %>% pull(shooting_time) %>% mean()
  shooting_time_s_shot1 = shots_athletes %>% filter(athlete_id == athletes$id[i] & mode == "S" & shot_number_series == 1) %>% pull(shooting_time) %>% mean()
  shooting_time_s = shots_athletes %>% filter(athlete_id == athletes$id[i] & mode == "S" & shot_number_series != 1) %>% pull(shooting_time) %>% mean()
  
  athletes_data[nrow(athletes_data) + 1,] = c(athletes$id[i], athletes$name[i], athletes$gender[i], athletes$nation[i],
                                              shots_athletes$pre_hit_rate_10[last_shot_of_athlete],
                                              shots_athletes$pre_hit_rate_10_mode[last_shot_of_athlete_p],
                                              shots_athletes$pre_hit_rate_10_mode[last_shot_of_athlete_s],
                                              shots_athletes$pre_hit_rate_50[last_shot_of_athlete],
                                              shots_athletes$pre_hit_rate_50_mode[last_shot_of_athlete_p],
                                              shots_athletes$pre_hit_rate_50_mode[last_shot_of_athlete_s],
                                              shots_athletes$pre_hit_rate_200[last_shot_of_athlete],
                                              shots_athletes$pre_hit_rate_200_mode[last_shot_of_athlete_p],
                                              shots_athletes$pre_hit_rate_200_mode[last_shot_of_athlete_s],
                                              shots_athletes$pre_hit_rate_200_mode_shotNr[last_shot_of_athlete_p_1],
                                              shots_athletes$pre_hit_rate_200_mode_shotNr[last_shot_of_athlete_p_2],
                                              shots_athletes$pre_hit_rate_200_mode_shotNr[last_shot_of_athlete_p_3],
                                              shots_athletes$pre_hit_rate_200_mode_shotNr[last_shot_of_athlete_p_4],
                                              shots_athletes$pre_hit_rate_200_mode_shotNr[last_shot_of_athlete_p],
                                              shots_athletes$pre_hit_rate_200_mode_shotNr[last_shot_of_athlete_s_1],
                                              shots_athletes$pre_hit_rate_200_mode_shotNr[last_shot_of_athlete_s_2],
                                              shots_athletes$pre_hit_rate_200_mode_shotNr[last_shot_of_athlete_s_3],
                                              shots_athletes$pre_hit_rate_200_mode_shotNr[last_shot_of_athlete_s_4],
                                              shots_athletes$pre_hit_rate_200_mode_shotNr[last_shot_of_athlete_s],
                                              shooting_time_p_shot1, shooting_time_p,
                                              shooting_time_s_shot1, shooting_time_s
  )
}

athletes_data <- athletes_data[order(athletes_data$name),] # TODO check warum NAs in athlete_data
save(athletes_data, file="shiny/biathlon_pred/athletes_data.RData")

athletes = add_athletes(shooting_series$Name,"m")
summary(athletes)