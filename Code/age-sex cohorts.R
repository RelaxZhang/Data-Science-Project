library(reshape2)

###Preprocess data###
SA3 <- read.csv("above1000.csv")
SA3[order(SA3$SA3.Code,SA3$Year),]
View(SA3[order(SA3$SA3.Code,SA3$Year),])
View(t(SA3[order(SA3$SA3.Code,SA3$Year),]))

sa3melt <- melt(SA3, id.vars=c(1:10))
ordered <- sa3melt[order(sa3melt$variable,sa3melt$SA3.Code, sa3melt$Year),]
ordered[,c(2,9,11,12)]
write.csv(ordered[,c(2,9,11,12)], "sa3melt.csv")
### ###



###Data viz###
#351 SA3s
#36 cohorts per SA3
library(forecTheta)

predictions <- c()
residuals <- c()

#index i indicates the SA3-cohort
#so for 351 SA3s with 36 cohorts each thats 12636 time series
for (i in 1:11700){
  indices <- (i-1)*21 + (1:16)
  currentseries <- as.ts(ordered[indices,c(2,10,11,12)][,4])
  currentforecast <- otm.arxiv(currentseries, thetaList = seq(1,5,by=0.5), 
                              approach = NULL, n1=6, m=5, H=5, p=2)
  predictions <- c(predictions,rep(-999,16), currentforecast$mean)
}

# sampleacf <- acf(ordered[1:21,c(2,10,11,12)][,4], lag.max = 20, plot=TRUE)
# cbind(sampleacf$lag,sampleacf$acf, sampleacf$snames)
# subsetpredicted <- ordered[1:3780,]
# subsetpredicted$fit <- predictions
# subsetpredicted$residual <- rep(NULL, dim(subsetpredicted)[1])
# subsetresiduals <- subsetpredicted[subsetpredicted$fit >= 0,]
# subsetresiduals[,c(2,10:13)]
# subsetresiduals$residuals <- subsetresiduals$fit-subsetresiduals$value
# subsetresiduals$ape <- 100*abs(subsetresiduals$residuals)/subsetresiduals$value
# View(subsetresiduals[,c(2,10:15)])
# View(subsetresiduals[subsetresiduals$Year == 2011,c(2,10:15)])

predicted <- ordered[1:245700,]
predicted$fit <- predictions
predictedresiduals <- predicted[predicted$fit >= 0,]
predictedresiduals$residuals <- predictedresiduals$fit- predictedresiduals$value
predictedresiduals$ape <- 100 * abs(predictedresiduals$residuals)/predictedresiduals$value


View(predictedresiduals[,c(2,10:15)])
View(predictedresiduals[predictedresiduals$ape>100,c(2,10:15)])


View(predictedresiduals[predictedresiduals$Year == 2011,c(2,10:15)])
View(predictedresiduals[predictedresiduals$ape>100 & predictedresiduals$Year == 2011,c(2,10:15)])
mean(predictedresiduals[predictedresiduals$Year == 2011,]$ape)

