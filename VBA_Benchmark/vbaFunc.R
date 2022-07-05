#################### Function Package of VBA to RStudio ########################

# Function for generating 3-Dimension jumpoffERP Array
jumpoffERP_Func <- function(agesex, ages, areas, year_male, year_female){
  jumpoffERP <- array(rep(0, areas * 2 * ages), dim=c(areas, 2, ages))
  row = 0
  for (i in 1:areas){
    for (a in 1:ages){
      row = row + 1
      jumpoffERP[i, 1, a] = agesex[row, year_female]
      jumpoffERP[i, 2, a] = agesex[row, year_male]}}
  return (jumpoffERP)
}

# Function for generating 1-Dimension xTFR Array
xTFR <- function(jumpoffERP, numareas){
  
  result = c(rep(0, numareas))
  
  for (i in 1:numareas){
    # Related age_sex group
    male04 = jumpoffERP[i, 2, 1]
    female04 = jumpoffERP[i, 1, 1]
    female_1519 = jumpoffERP[i, 1, 4]
    female_2024 = jumpoffERP[i, 1, 5]
    female_2529 = jumpoffERP[i, 1, 6]
    female_3034 = jumpoffERP[i, 1, 7]
    female_3539 = jumpoffERP[i, 1, 8]
    female_4044 = jumpoffERP[i, 1, 9]
    female_4549 = jumpoffERP[i, 1, 10]
  
    # Subjective definition for simplifying formula
    new_born = male04 + female04
    most_fert = female_2529 + female_3034
    all_fert = (female_1519 + female_2024 + female_2529 + female_3034 +
                female_3539 + female_4044 + female_4549)
    
    # Calculate result for xTFR in each area
    result[i] = (10.65 - (12.55 * most_fert / all_fert)) * new_born / all_fert
  }
  # Return xTRF result for storing into Fertility Sheet
  return (result)
}

################################################################################
# Function for recording ERP in small area with age & sex information
readERP <- function(intervals, numareas, numages, agesex, project){
  
  # Create empty array for storing ERP in small area by age and sex
  infoERP <- array(rep(0, intervals * numareas * 2 * numages),
                   dim=c(intervals, numareas, 2, numages))

  # Collect data
  for (year in 1:intervals){
    row = 0
    for (i in 1:numareas){
      for (a in 1:numages){
        row = row + 1
        
        # Identify Current Process (Input / Projection)
        if (project){
          index = year + 1
        } else{
          index = year}
        
        # Females
        infoERP[year, i, 1, a] = agesex[row, (6 + index)]
        # Males
        infoERP[year, i, 2, a] = agesex[row, (4 + index)]}}}

  # Return ERP data with given year range(s)
  return(infoERP)
}

# Function to load Life Expectancy at Birth Assumptions, Mortality Surface Data
readLEMS <- function(final, numareas, numages, nLxMS, expectancy, mortality){
  
  # Create empty array for storing Life expectancy and Mortality Surface
  eO <- array(rep(0, (final+1) * numareas * 2), dim=c((final+1), numareas, 2))
  MS_nLx <- array(rep(0, nLxMS * 2 * numages), dim=c(nLxMS, 2, numages))
  
  # Collect Life Expectancy Data
  for (i in 1:numareas){
    col_eO = 0
    for (a in 1:(final + 1)){
      col_eO = col_eO + 1
      eO[a, i, 1] = expectancy[1, col_eO]
      eO[a, i, 2] = expectancy[1, col_eO + 4]}}

  # Collect Mortality Surface Data
  for (i in 1:(numages)){
    col_MS = 2
    for (a in 1:(nLxMS)){
      col_MS = col_MS + 1
      MS_nLx[a, 1, i] = mortality[i, col_MS]
      MS_nLx[a, 2, i] = mortality[i + numages + 1, col_MS]}}
    
  return (list("eO" = eO, "MS_nLx" = MS_nLx))
}

################################################################################