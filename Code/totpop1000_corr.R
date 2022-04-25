# Correlation Calculation
totpop_1000 <- read.csv(file="../Data/totpop_1000.csv")
corr_matrix <- as.matrix(cor(totpop_1000))

# Extract regions for presentation
sample_corr <- corr_matrix[c("Latrobe Valley", "Barossa", "Wellington", "Gawler - Two Wells"),
                           c("Latrobe Valley", "Barossa", "Wellington", "Gawler - Two Wells")]

# Save the correlation matrix as csv for further analysis
write.csv(data.frame(corr_matrix), paste0("totpop1000_corrmatrix.csv"))