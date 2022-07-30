SA3 <- read.csv("above1000.csv")

sa3codes <- unique(SA3$SA3.Code)
newSA3 <- SA3

for (i in 11:46){ #column indices for the cohorts
  for (sa3 in sa3codes){
    name <- paste(sa3, colnames(SA3)[i]) #creates string "{SA3Code} {cohort}"
    newSA3[[name]] <- SA3[SA3$SA3.Code == sa3, i]
  }
}

newSA3 <- newSA3[1:21, c(2,48:dim(newSA3)[2])] #only include year and new columns
newSA3$Year <- 1991:2011

write.csv(newSA3,"newSA3.csv")