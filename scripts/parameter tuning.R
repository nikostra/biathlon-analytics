# This script was used to perform the CV hyperparameter search to find the best model
# requisite loaded shots data

# Set the seed for reproducibility
set.seed(123)

# prepare shots by dropping NA values and redundant variables (race ID, athlete ID)
shots_prepared = shots %>% drop_na() %>% select(-ends_with('id'))

# Create an index for the train/test split
index <- createDataPartition(shots_prepared$target, p = 0.9, list = FALSE)

# Create the training set
train_set <- shots_prepared[index, ]
sparse_matrix_train <- sparse.model.matrix(target ~ ., data = train_set)[,-1]
y_train = as.factor(train_set[,1])

# Create the test set
test_set <- shots_prepared[-index, ]
y_test = test_set$target
sparse_matrix_test <- sparse.model.matrix(target ~ ., data = test_set)[,-1]

# create the grid for the hyperparameter search
param_grid <- expand.grid(max_depth = c(4, 3, 6),
                          eta = c(0.1, 0.05, 0.15, 0.2),
                          nrounds = c(100, 150, 300, 400),
                          gamma = 0,
                          colsample_bytree = 1,
                          min_child_weight = 1,
                          subsample = 1)

control <- trainControl(method = "cv", number = 10)

# perform CV hyperparameter search
xgb_model <- train(x = sparse_matrix_train,
                   y = y_train,
                   trControl = control,
                   method = "xgbTree",
                   metric = "Accuracy",
                   tuneGrid = param_grid,
                   verbosity = 0)

# show result 
print(xgb_model$bestModel)
print(xgb_model$bestTune)
summary(xgb_model)

# save a model with the best config using the full data
# will be used in the shiny web app
fullData = sparse.model.matrix(target ~ ., data = shots_prepared)[,-1]
y = shots_prepared$target
saveModel = xgboost::xgboost(data=fullData,label = y, objective = "binary:logistic",
                    nrounds = 150, max_depth=3,eta=0.15)
xgb.save(saveModel, 'xgb.model')
model = xgboost::xgb.load("shiny/biathlon_pred/xgb.model")

# get AUROC metric for best model using test data
predictions = (predict(model, sparse_matrix_test))
roc_data <- roc(y_test, predictions)
auc(roc_data)

best_auc = auc(roc_data)
plot(roc_data, main = "ROC Curve")
# best roc = 0.6469
# treshhold according to ROC = 0.8072

# view feature importance
importanceRaw <- xgb.importance(feature_names = colnames(sparse_matrix_train), model = model)
importanceClean <- importanceRaw[,`:=`(Cover=NULL, Frequency=NULL)]

head(importanceClean, n=10)
xgb.plot.importance(importance_matrix = importanceClean)

# save example matrix for use in the shiny web aopp
test_na = as.data.frame(test_set[c(which.min(test_set$wind_speed),which.max(test_set$wind_speed)),])
test_na$pre_hit_rate_200_mode = NA
sparse_matrix_test2 <- sparse_matrix_test[c(1,2,3,4,5),]
sparse_matrix_test2[1,45] = NA
sparse_matrix_test2[1,26] = 0
model_matrix = sparse_matrix_test2
save(model_matrix, file="shiny/biathlon_pred/model_matrix.RData")
predict(xgb_model$finalModel, sparse_matrix_test2)

