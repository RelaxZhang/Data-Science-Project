# Import the libraries
library("reshape")
library("ggplot2")
library("plot.matrix")
library("MASS")

# Read in dataset of SA3 regions with total population greater than 1000 in each year
above1000 <- read.csv(file="../Data/ts_1000.csv")

# Aggregate data with sex type in each region
# Hierarchy becomes: Total -> Male & Female among each year in each region
sex <- aggregate(cbind(above1000$X1991, above1000$X1992, above1000$X1993, 
                       above1000$X1994, above1000$X1995, above1000$X1996, 
                       above1000$X1997, above1000$X1998, above1000$X1999, 
                       above1000$X2000, above1000$X2001, above1000$X2002,
                       above1000$X2003, above1000$X2004, above1000$X2005,
                       above1000$X2006, above1000$X2007, above1000$X2008,
                       above1000$X2009, above1000$X2010, above1000$X2011),
                 by = list(above1000$sex, above1000$SA3_Code), FUN = sum)

aggre_sex <- c("sex", "SA3", 1991, 1992, 1993, 1994, 1995, 1996, 1997, 1998, 1999, 2000,
               2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011)
colnames(sex) <- aggre_sex

# Aggregate data with age type in each region
# Hierarchy becomes: Total -> 0-4 & 5-9 ... & 85+ among each year in each region
age <- aggregate(cbind(above1000$X1991, above1000$X1992, above1000$X1993, 
                       above1000$X1994, above1000$X1995, above1000$X1996, 
                       above1000$X1997, above1000$X1998, above1000$X1999, 
                       above1000$X2000, above1000$X2001, above1000$X2002,
                       above1000$X2003, above1000$X2004, above1000$X2005,
                       above1000$X2006, above1000$X2007, above1000$X2008,
                       above1000$X2009, above1000$X2010, above1000$X2011),
                 by = list(above1000$age, above1000$SA3_Code), FUN = sum)

aggre_age <- c("age", "SA3", 1991, 1992, 1993, 1994, 1995, 1996, 1997, 1998, 1999, 2000,
               2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011)

timestamp <- c(1991, 1992, 1993, 1994, 1995, 1996, 1997, 1998, 1999, 2000, 2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011)

colnames(age) <- aggre_age

# Collect Total population among each year in each SA3 region
total <- c()
drop_sex <- c()
drop_age <- c()

for (i in (1:length(sex[, 1]))){
  if (sex[i, 1] == "T"){
    total = rbind(total, (sex[i, ]))
    drop_sex = append(drop_sex, i)}}

for (i in (1:length(age[, 1]))){
  if (age[i, 1] == 99){
    drop_age = append(drop_age, i)}}

# Drop total rows in both aggregated age and sex dataset
sex <- sex[-drop_sex, ]
age <- age[-drop_age, ]

f_list <- seq(from = 1, to = 650, by = 2)
m_list <- seq(from = 2, to = 650, by = 2)

for (i in length(1:325)){
  sample_sex_f <- t(sex[f_list[i], ][3:23])
  sample_sex_m <- t(sex[m_list[i], ][3:23])
  sample_sex_t <- t(total[i, ][3:23])
  plot_data <- data.frame(cbind(sample_sex_f, sample_sex_m, sample_sex_t, timestamp))
  plot_data_test <- melt(plot_data, id.vars = "timestamp")
  ggplot(data = plot_data_test, aes(x = timestamp, y = value, color = variable)) + geom_line()
}

age_sample <- age[1:18, 3:23]

write.csv(data.frame(sex), paste0("../Data/sex_1000.csv"))
write.csv(data.frame(age), paste0("../Data/age_1000.csv"))
write.csv(data.frame(total), paste0("../Data/tot_1000.csv"))