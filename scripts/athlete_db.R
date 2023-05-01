athletes = data.frame(id=c(1), name=c("wright campbell"), gender=c("m"))

# names is array with all names, gender is string "m" or "f"
add_athletes = function(names, gender){
  for (name in names){
    name = tolower(name)
    if(!(name %in% athletes$name)){
      athletes[nrow(athletes) + 1,]$id = as.numeric(max(athletes$id)) + 1
      athletes[nrow(athletes),]$name= name
      athletes[nrow(athletes),]$gender = gender
      }
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



athletes = add_athletes(shooting_series$Name,"m")
summary(athletes)