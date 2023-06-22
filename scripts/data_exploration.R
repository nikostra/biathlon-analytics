library(tidyverse)
library(xgboost)
library(data.table)
library(Matrix)
library(pROC)
library(explore)
library(openxlsx)
library(tidyxl)
library(readxl)
library(data.table)
library(caret)
library(rpart)
library(rpart.plot)

athletes = data.frame(id=c(1), name=c("wright campbell"), gender=c("m"), nation = c("NZL"))
race_id = 1
shots = load_shots()
shots_short = load_shots(FALSE)
shots_short = shots_short[,-4]
summary(shots)
data = shots

y = (shots$target)

# try regression
regData = shots[,2:ncol(shots)]
regData = regData[,-3]

model = glm(y ~ ., data = regData)
summary(model)
min(na.omit(predict(model, shots)))

# xgboost
sparse_matrix <- sparse.model.matrix(target ~ ., data = shots)[,-1]
sparse_matrix = sparse_matrix[,-c(11,22)] # remove athlete and race ID
model <- xgboost(data = sparse_matrix, label = y, max_depth = 12,
               eta = 0.3, nthread = 10, nrounds = 300,objective = "binary:logistic")
importance <- xgb.importance(feature_names = colnames(sparse_matrix), model = model)
head(importance, n = 10)

# split in test and train
train = shots %>% filter(season != "1819") %>% drop_na() %>% select(-ends_with('id'))
sparse_matrix_train <- sparse.model.matrix(target ~ ., data = train)[,-1]
y_train = train[,1]

test = shots %>% filter(season == "1819") %>% drop_na() %>% select(-ends_with('id'))
y_test = test$target
sparse_matrix_test <- sparse.model.matrix(target ~ ., data = test)[,-1]

model <- xgboost(data = sparse_matrix_train, label = y_train, max_depth = 3,
                 eta = 0.3, nthread = 16, nrounds = 100,objective = "binary:logistic")

predictions = (predict(model, sparse_matrix_test))

# feature importance
importanceRaw <- xgb.importance(feature_names = colnames(sparse_matrix_train), model = model)
importanceClean <- importanceRaw[,`:=`(Cover=NULL, Frequency=NULL)]

head(importanceClean)
xgb.plot.importance(importance_matrix = importanceClean)

#R0C
roc_data <- roc(y_test, predictions)
auc(roc_data)
# Plot the ROC curve
plot(roc_data, main = "ROC Curve")

# data exploration
shots_split = shots_split %>% mutate(target = shots$target)
explore_targetpct(shots, wind_cat) + xlab("Category of wind speed")
explore_targetpct(shots %>% filter(mode == "S" & shot_number_series == 5), 
                  last_shot) + xlab("Last shot")
explore_targetpct(shots, competition_level)

#rpart
tree_data = shots
tree_data$target = as.factor(tree_data$target)
tree = rpart(target ~ ., data = tree_data, parms = list(split = "information"), cp = 0.000075)
tree
fancyRpartPlot(tree, caption = NULL)
