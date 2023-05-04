library(tidyverse)
library(xgboost)
library(data.table)
library(Matrix)
library(pROC)

athletes = data.frame(id=c(1), name=c("wright campbell"), gender=c("m"))
race_id = 1
shots = load_shots()
shots_short = load_shots(FALSE)
summary(shots)
data = shots

y = (shots$target)

# try regression
regData = shots[,2:21]

model = glm(y ~ ., data = regData)
summary(model)
min(na.omit(predict(model, shots)))

# xgboost
sparse_matrix <- sparse.model.matrix(target ~ ., data = shots)[,-1]
sparse_matrix = sparse_matrix[,-c(11,22)] # remove athlete and race ID
model <- xgboost(data = sparse_matrix, label = y, max_depth = 12,
               eta = 0.3, nthread = 2, nrounds = 20,objective = "binary:logistic")
importance <- xgb.importance(feature_names = colnames(sparse_matrix), model = model)
head(importance, n = 10)

train = shots %>% filter(season != "2122")
y_train = train$target
test = shots %>% filter(season == "2122")
y_test = test$target
sparse_matrix_train <- sparse.model.matrix(target ~ ., data = train)[,-1]
sparse_matrix_train = sparse_matrix_train[,-c(11,22)] # remove athlete and race ID
sparse_matrix_test <- sparse.model.matrix(target ~ ., data = test)[,-1]
sparse_matrix_test = sparse_matrix_test[,-c(11,22)] # remove athlete and race ID

model <- xgboost(data = sparse_matrix_train, label = y_train, max_depth = 12,
                 eta = 0.5, nthread = 2, nrounds = 30,objective = "binary:logistic")

predictions = (predict(model, sparse_matrix_test))

# feature importance
importanceRaw <- xgb.importance(feature_names = colnames(sparse_matrix_train), model = model)
importanceClean <- importanceRaw[,`:=`(Cover=NULL, Frequency=NULL)]

head(importanceClean)
xgb.plot.importance(importance_matrix = importance)

#R0C
roc_data <- roc(y_test, predictions)
auc(roc_data)
# Plot the ROC curve
plot(roc_data, main = "ROC Curve")
