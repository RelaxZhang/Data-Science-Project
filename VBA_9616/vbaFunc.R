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
# Function for Generating Input Data for ASFR (Area specific Fertility Rate)
inputASFR <- function(final, numareas, age_groups, prelimASFR, modelASFR, TFR, modelTFR){
  
  # Create empty array for storing ASFR input Data
  ASFR <- array(rep(0, (final+1) * numareas * age_groups), dim=c((final+1), numareas, age_groups))
  
  # Collect ASFR input data
  for (y in 1:(final + 1)){
    row = 0
    for (i in 1:numareas){
      row = row + 1
      if (prelimASFR[row, 1] != 0){
        prelimtot = 0
        for (a in 1:age_groups){
          prelimtot = prelimtot + prelimASFR[i, a]}
        for (a in 1:age_groups){
          ASFR[y, i, a] = prelimASFR[i, a] * TFR[y, i] / (prelimtot * 5)}}
      else{
        for (a in 1:age_groups){
          ASFR[y, i, a] = modelASFR[a] * TFR[y, i] / modelTFR}}}}

  # Return ASFR input Data
  return (ASFR)
}

# Function for estimating base periold birth by sex
inputbirth <- function(numareas, age_groups, ASFR_data, ERP, SRB){
  
  # Amount of Each Age group female could give birth to
  tempbirths = rep(0, age_groups)
  # Total estimate birth in different sex
  tempbirthssex = array(rep(0, numareas * 2), dim=c(numareas, 2))
  # Total estimate birth
  temptotbirths = rep(0, numareas)

  # Collect estimated total birth population and sex birth population
  for (i in 1:numareas){
    temptotbirths[i] = 0

    for (a in 1:age_groups){
      tempbirths[a] = ASFR_data[1, i, a] * 2.5 * (ERP[1, i, 1, a + 3] + ERP[2, i, 1, a + 3])
      temptotbirths[i] = temptotbirths[i] + tempbirths[a]}

    tempbirthssex[i, 1] = temptotbirths[i] * (100 / (SRB + 100))
    tempbirthssex[i, 2] = temptotbirths[i] * (SRB / (SRB + 100))}

  # Return tempbirthssex, temptotbirths estimation data
  return (list("tempsex" = tempbirthssex, "temp" = temptotbirths))
}

# Function for generating ASDR Input Data for further usage
inputASDR <- function(final, numareas, lastage, nLxMS, eO, MS_TO, MS_nLx){
  
  # Create empty array for storing ASDR input data
  ASDR = array(rep(0, (final+1) * numareas * 2 * (lastage+1)), dim=c((final+1), numareas, 2, (lastage+1)))

  for (i in 1:numareas){
    Lower_TO = array(rep(0, (final + 1) * 2), dim=c((final + 1), 2))
    Upper_TO = array(rep(0, (final + 1) * 2), dim=c((final + 1), 2))
    Lower_num = array(rep(0, (final + 1) * 2), dim=c((final + 1), 2))
    Upper_num = array(rep(0, (final + 1) * 2), dim=c((final + 1), 2))
    proportion = array(rep(0, (final + 1) * 2), dim=c((final + 1), 2))

    # Find out where in the mortality surface each eO lies
    for (y in 1:(final + 1)){
      for (s in 1:2){
        for (z in 1:(nLxMS - 1)){
          e0_100k = eO[y, i, s] * 100000
          if (e0_100k >= MS_TO[z, s] & e0_100k < MS_TO[z + 1, s]){
            Upper_TO[y, s] = MS_TO[z + 1, s]
            Upper_num[y, s] = z + 1
            Lower_TO[y, s] = MS_TO[z, s]
            Lower_num[y, s] = z
            proportion[y, s] = (e0_100k - Lower_TO[y, s]) / (Upper_TO[y, s] - Lower_TO[y, s])
            break}}}}

    # Create nLx values for each eO
    nLx = array(rep(0, (final + 1) * 2 * (lastage + 1)), dim=c((final + 1), 2, (lastage + 1)))
    for (y in 1:(final + 1)){
      for (s in 1:2){
        lower = 0 # Ensure there is lower value without the above loop
        upper = 0 # Ensure there is upper value without the above loop
        lower = Lower_num[y, s]
        upper = Upper_num[y, s]
        for (a in 1:(lastage + 1)){
          MS_low = MS_nLx[lower, s, a]
          MS_up = MS_nLx[upper, s, a]
          nLx[y, s, a] = MS_low + proportion[y, s] * (MS_up - MS_low)}}}

    # Create Period-Cohort ASDRs (Area Specific Death Rates)
    for (y in 1:(final + 1)){
      for (s in 1:2){
      
        # Record Area Specific Death Rate for age groups excluding 0-4 and 85+
        for (pc in 2:lastage){
          younger = nLx[y, s, pc-1]
          elder = nLx[y, s, pc]
          ASDR[y, i, s, pc] = (younger - elder) / (5 / 2 * (younger + elder))}

        # Record ASDR for 0-4
        ASDR[y, i, s, 1] = ((5 * 100000) - nLx[y, s, 1]) / (5 / 2 * nLx[y, s, 1])
        # Record ASDR for 85+
        younger = nLx[y, s, lastage]
        elder = nLx[y, s, lastage+1]
        ASDR[y, i, s, lastage+1] = ((younger + elder) - elder) / (5 / 2 * (younger + elder + elder))}}}

  # Return Area Specific Death Rates Input Data
  return (ASDR)
}

# Function for calculating the estimated bas period deaths for the population accounts
inputdeath <- function(numareas, lastage, ASDR, ERP){
  
  # Number of death in each age-sex group
  tempdeaths = array(rep(0, numareas * 2 * (lastage + 1)), dim=c(numareas, 2, (lastage + 1)))
  # Total estimate death
  temptotdeaths = rep(0, numareas)

  # Record number of death in each age-sex cohort
  for (i in 1:numareas){
    for (s in 1:2){
      # Record estimated Deaths for age-sex groups excluding 0-4 and 85+
      for (pc in 2:lastage){
        tempdeaths[i, s, pc] = ASDR[1, i, s, pc] * 2.5 * (ERP[1, i, s, pc - 1] + ERP[2, i, s, pc])
      # Record Deaths for 0-4
      tempdeaths[i, s, 1] = ASDR[1, i, s, 1] * 2.5 * ERP[2, i, s, 1]
      # Record Deaths for 85+
      younger = ERP[1, i, s, lastage]
      elder_f = ERP[1, i, s, lastage + 1]
      elder_m = ERP[2, i, s, lastage + 1]
      tempdeaths[i, s, lastage + 1] = ASDR[1, i, s, lastage + 1] * 2.5 * (younger + elder_f + elder_m)
      }
    }
    # Record estimated total death
    temptotdeaths[i] = 0
    for (s in 1:2){
      for (pc in 1:(lastage + 1)){
        temptotdeaths[i] = temptotdeaths[i] + tempdeaths[i, s, pc]}}}

  # Return tempdeaths, temptotdeaths estimation data
  return (list("deaths" = tempdeaths, "totdeaths" = temptotdeaths))
}

# Function for calculating migration input data
inputMigration <- function(final, numareas, lastage, modelASMR, ERP, totmig, TotPop, temptotdeaths,
                           temptotbirths, tempdeaths, tempbirthssex, ASDR, Areaname, sexlabel, pclabel, intervallabel){
  
  # Create empty array for storing migration related data
  # Area Specific Out Migration Ratio
  ASOMR = array(rep(0, (final+1) * numareas * 2 * (lastage+1)), dim=c((final+1), numareas, 2, (lastage+1)))
  # Inward migration amount
  inward = array(rep(0, (final+1) * numareas * 2 * (lastage+1)), dim=c((final+1), numareas, 2, (lastage+1)))
  # Outward migration amount
  outward = array(rep(0, numareas * 2 * (lastage+1)), dim=c(numareas, 2, (lastage+1)))
  # Prelimium migration
  prelimmig = array(rep(0, numareas * 2 * (lastage+1)), dim=c(numareas, 2, (lastage+1)))
  # Scaled migration
  scaledmig = array(rep(0, numareas * 2 * (lastage+1)), dim=c(numareas, 2, (lastage+1)))
  # Prelimium inward migration
  prelimbaseinward = array(rep(0, numareas * 2 * (lastage+1)), dim=c(numareas, 2, (lastage+1)))
  # Prelimium outward migration
  prelimbaseoutward = array(rep(0, numareas * 2 * (lastage+1)), dim=c(numareas, 2, (lastage+1)))
  # Cohort net migration
  cohortnetmig = array(rep(0, numareas * 2 * (lastage+1)), dim=c(numareas, 2, (lastage+1)))
  # Adjuested net migration
  adjnetmig = array(rep(0, numareas * 2 * (lastage+1)), dim=c(numareas, 2, (lastage+1)))
  # Raw scaling factor
  scalingfactor2a = array(rep(0, numareas * 2 * (lastage+1)), dim=c(numareas, 2, (lastage+1)))
  # Smoothed scaling factor
  scalingfactor2b = array(rep(0, numareas * 2 * (lastage+1)), dim=c(numareas, 2, (lastage+1)))
  
  # Collecting Migration Input Data
  for (i in 1:numareas){
    
    # Generate preliminary migration turnover by sex-age cohort based on model rates
    for (s in 1:2){
      # Generate preliminary migration turnover for all sex-age cohorts exclude 0-4, 85+
      for (pc in 2:lastage){
        prelimmig[i, s, pc] = modelASMR[s, pc] * 2.5 * (ERP[1, i, s, pc - 1] + ERP[2, i, s, pc])
  
      # Generate preliminary migration turnover for 0-4
      prelimmig[i, s, 1] = modelASMR[s, 1] * 2.5 * ERP[2, i, s, 1]
      
      # Generate preliminary migration turnover for 85+
      prelimmig[i, s, lastage+1] = modelASMR[s, lastage+1] * 2.5 * (ERP[1, i, s, lastage] + ERP[1, i, s, lastage+1] + ERP[2, i, s, lastage+1])}}

    # Generate total preliminarymigration turnover value
    totprelimmig = 0
    for (s in 1:2){
      for (pc in 1:(lastage + 1)){
        totprelimmig = totprelimmig + prelimmig[i, s, pc]}}

    # Collecting scaled preliminary migration by sex-age cohort to migration turnover 
    # as calculated by the crude migration turnover rate; and then divide by 2 to give the initial inward and outward migration
    for (s in 1:2){
      for (pc in 1:(lastage + 1)){
        scaledmig[i, s, pc] = prelimmig[i, s, pc] * (0.5 * totmig[i]) / totprelimmig}}

    # Calculate residual net migration for base period in each area i
    basenetmig = TotPop[2, i] - TotPop[1, i] + temptotdeaths[i] - temptotbirths[i]

    # Calculate the scaling factor to estimate separate inward and outward migration totals which are consistent with residual net migration
    scalingfactor1 = (basenetmig + (basenetmig ^ 2 + (4 * 0.5 * totmig[i] * 0.5 * totmig[i])) ^ 0.5) / (2 * 0.5 * totmig[i])

    # Calculate preliminary inward and outward migration by sex-age cohort (Different with the Original Value in 'Account')
    for (s in 1:2){
      for (pc in 1:(lastage + 1)){
        prelimbaseinward[i, s, pc] = scaledmig[i, s, pc] * scalingfactor1
        prelimbaseoutward[i, s, pc] = scaledmig[i, s, pc] / scalingfactor1}}

    # Calculate cohort-specific residual net migration (for adjustment usage)
    for (s in 1:2){
      # Generate cohort-specific residual net migration for all sex-age cohorts exclude 0-4, 85+
      for (pc in 2:lastage){
        cohortnetmig[i, s, pc] = ERP[2, i, s, pc] - ERP[1, i, s, pc - 1] + tempdeaths[i, s, pc]}

      # Generate cohort-specific residual net migration for 0-4
      cohortnetmig[i, s, 1] = ERP[2, i, s, 1] - tempbirthssex[i, s] + tempdeaths[i, s, 1]
      
      # Generate cohort-specific residual net migration for 85+
      cohortnetmig[i, s, lastage + 1] = ERP[2, i, s, lastage + 1] - (ERP[1, i, s, lastage] + ERP[1, i, s, lastage + 1]) + tempdeaths[i, s, lastage + 1]}

    # Adjusted cohort-specific net migration with averaged over sex for selected ages
    for (s in 1:2){
      for (pc in 1:14){
        adjnetmig[i, s, pc] = 0.5 * (cohortnetmig[i, 1, pc] + cohortnetmig[i, 2, pc])}
      for (pc in 15:(lastage + 1)){
        adjnetmig[i, s, pc] = cohortnetmig[i, s, pc]}}

    # Calculate scaling factors to adjust directional migration (Raw & Smoothed)
    # (a) Raw Scaling factor
    for (s in 1:2){
      for (pc in 1:(lastage + 1)){
        if (prelimbaseinward[i, s, pc] > 0){
          adjmig = adjnetmig[i, s, pc]
          preinward = prelimbaseinward[i, s, pc]
          preoutward = prelimbaseoutward[i, s, pc]
          scalingfactor2a[i, s, pc] = (adjmig + (adjmig ^ 2 + (4 * preinward * preoutward)) ^ 0.5) / (2 * preinward)}
      else{
        scalingfactor2a[i, s, pc] = 1}}}

    # (b) Smoothed Scaling factor
    for (s in 1:2){
      for (pc in 1:(lastage + 1)){
        if (pc < 11){
          scalingfactor2b[i, s, pc] = scalingfactor2a[i, s, pc]}
        else if (pc == (lastage + 1)){
          scalingfactor2b[i, s, pc] = 1 / 2 * (scalingfactor2a[i, s, pc - 1] + scalingfactor2a[i, s, pc])}
        else{
          scalingfactor2b[i, s, pc] = 1 / 3 * (scalingfactor2a[i, s, pc - 1] + scalingfactor2a[i, s, pc] + scalingfactor2a[i, s, pc + 1])}

        # Further guarantee of no too extreme scaling factor value
        if (scalingfactor2b[i, s, pc] > 10){
          scalingfactor2b[i, s, pc] = 10}
        else if(scalingfactor2b[i, s, pc] < 0.1){
          scalingfactor2b[i, s, pc] = 0.1}}}

        # Adjust inward and outward migration with smoothed scaling factor
        for (s in 1:2){
          for (pc in 1:(lastage + 1)){
            inward[1, i, s, pc] = prelimbaseinward[i, s, pc] * scalingfactor2b[i, s, pc]
            outward[i, s, pc] = prelimbaseoutward[i, s, pc] / scalingfactor2b[i, s, pc]}}

        # Smoothing inward migration age profiles at selected age groups
        intemp = array(rep(0, 2 * (lastage+1)), dim=c(2, (lastage+1)))
        for (s in 1:2){
          for (pc in 1:(lastage + 1)){
            intemp[s, pc] = inward[1, i, s, pc]}}
        for (s in 1:2){
          for (pc in 1:(lastage - 1)){
            inward[1, i, s, pc] = intemp[s, pc]}
          for (pc in lastage:(lastage + 1)){
            inward[1, i, s, pc] = 1 / 2 * (intemp[1, pc] + intemp[2, pc])}}

        # Calculate ASOMR (Area Specific Out Migration Rate)
        for (s in 1:2){
          # Record ASOMR input for all sex-age cohort exclude 0-4 & 85+
          for (pc in 1:(lastage-1)){
            if (ERP[1, i, s, pc] + ERP[2, i, s, pc+1] > 0){
              ASOMR[1, i, s, pc+1] = outward[i, s, pc+1] / (2.5 * (ERP[1, i, s, pc] + ERP[2, i, s, pc+1]))}
            else{
              ASOMR[1, i, s, pc+1] = 0
              print("Warning: Base period out-migration rate canoot be calculated due to zero population")
              print("Aera = ", Areaname[i], " & Sex = ", sexlabel[s], " & Period-Cohort = ", pclabel[pc+1])}}

          # Record ASOMR input for sex-age cohort 0-4
          if (ERP[2, i, s, 1] > 0){
            ASOMR[1, i, s, 1] = outward[i, s, 1] / (2.5 * ERP[2, i, s, 1])}
          else{
            ASOMR[1, i, s, 1] = 0
            print("Warning: Base period out-migration rate canoot be calculated due to zero population")
            print("Aera = ", Areaname[i], " & Sex = ", sexlabel[s], " & Period-Cohort = ", pclabel[1])}

          # Record ASOMR inpurt for sex age ohort 85+
          if (ERP[1, i, s, lastage] + ERP[1, i, s, lastage+1] > 0){
            ASOMR[1, i, s, lastage+1] = outward[i, s, lastage+1] / (2.5 * (ERP[1, i, s, lastage] + ERP[1, i, s, lastage+1] + ERP[2, i, s, lastage+1]))}
          else{
            ASOMR[1, i, s, lastage+1] = 0}}

        # Smoothing ASOMR at selected age groups
        outtemp = array(rep(0, 2 * (lastage+1)), dim=c(2, (lastage+1)))
        for (s in 1:2){
          for (pc in 1:(lastage + 1)){
            outtemp[s, pc] = ASOMR[1, i, s, pc]}}
        # Start smoothing
        for (s in 1:2){
          # When for the first 11 groups, keep ASOMR unchanged
          for (pc in 1:11){
            ASOMR[1, i, s, pc] = outtemp[s, pc]}

          # Smooth the rest of the elder age groups ASOMR
          for (pc in 12:(lastage + 1)){
            ASOMR[1, i, s, pc] = 1 / 2 * (outtemp[1, pc] + outtemp[2, pc])}}

        # Set default ASOMR and Inward migration projection on the future 3 year intervals (20-25, 25-30, 40-35)
        for (y in 2:(final + 1)){
          for (s in 1:2){
            for (pc in 1:(lastage + 1)){
              ASOMR[y, i, s, pc] = ASOMR[1, i, s, pc]
              inward[y, i, s, pc] = inward[1, i, s, pc]}}}

        # Check that ASOMRs + ASDR do not exceed 0.4 & Adjust ASOMRs if necessary
        # Helps to avoid negative population if there is no inward migration because pop-at-risk = 2.5 * (Pop0 + Pop1)
        for (y in 1:(final + 1)){
          for (s in 1:2){
            for (pc in 1:(lastage + 1)){
              if (ASOMR[y, i, s, pc] + ASDR[y, i, s, pc] > 0.5){
                temprate = ASOMR[y, i, s, pc]
                ASOMR[y, i, s, pc] = 0.4 - ASDR[y, i, s, pc]
                print("Warning: Excessive outward migration rate reducing")
                print("Avoid having outward migration rate + age-specific death rate from exceeding 0.4")
                print("Old rate = ", temprate, " & New rate = ", ASOMR[y, i, s, pc], " & Area = ", Areaname[i], " & Sex = ",
                      sexlabel[s], " & Period-Cohort = ", pclabel[pc], " & Interval = ", intervallabel[y])}}}}}

  # Return migration input data
  return (list("ASOMR" = ASOMR, "inward" = inward, "outward" = outward, "prelimmig" = prelimmig, "scaledmig" = scaledmig,
               "prelimbaseinward" = prelimbaseinward, "prelimbaseoutward" = prelimbaseoutward, "cohortnetmig" = cohortnetmig,
               "adjnetmig" = adjnetmig, "scalingfactor2a" = scalingfactor2a, "scalingfactor2b" = scalingfactor2b))
}

################################################################################
# Function for read in National Projection Data
readNatP <- function(final, numages, TotPopulation){
  
  # Create empty array for storing national total population
  NatP = array(rep(0, (final + 1) * 2 * numages), dim = c((final + 1), 2, numages))
  
  # Load in Total Population Data
  row = 0
  for (s in 1:2){
    for (a in 1:numages){
      row = row + 1
      col = 3
      for (y in 1:(final + 1)){
        col = col + 1
        NatP[y, s, a] = TotPopulation[row, col]}}}
  
  # Return National Projection data
  return (NatP)
}

# Function for reading in national projected deaths, new migration data
readDN <- function(final, lastage, numages, dataset){
  
  # Create empty array for storing national population data (death / net migration)
  Natpc = array(rep(0, final * 2 * numages), dim = c(final, 2, numages))
  
  # Load in National Population Data
  row = 0
  for (s in 1:2){
    for (pc in 1:(lastage + 1)){
      row = row + 1
      col = 2
      for (y in 1:final){
        col = col + 1
        Natpc[y, s, pc] = dataset[row, col]}}}
  
  # Return National Projection Data (in Period-cohort)
  return (Natpc)
}

# Function for writing headings / row name in CheckMig, CheckDeaths, Log Sheets
writeNoteCL <- function(lastage, pclabel, labels_other_key, final, Areaname, sexlabel){
  
  # Create empty array for creating framework of check sheets
  wb_wt_Log = array(rep(NaN, final*3), dim = c(final, 3))
  wb_wt_CheckMig = array(rep(NaN, (final*(lastage + 1)*2) * (numareas * 2 + 3 + 1)),
                         dim = c(final*(lastage+1)*2, numareas * 2 + 3 + 1))
  wb_wt_CheckDeaths = array(rep(NaN, (final*(lastage + 1)*2) * (numareas + 3 + 1)),
                            dim = c(final*(lastage+1)*2, numareas + 3 + 1))
  
  # Load in Log Headings
  index = 1
  for (i in 1:final){
    index = index + 1
    wb_wt_Log[i] = labels_other_key[index, 2]}
  
  index = 0
  for (f in 1:final){
    for (s in 1:2){
      for (a in 1:(lastage + 1)){
        index = index + 1
        wb_wt_CheckMig[index, 1] = labels_other_key[f + 1, 2]
        wb_wt_CheckMig[index, 2] = sexlabel[s]
        wb_wt_CheckMig[index, 3] = pclabel[a]
        
        wb_wt_CheckDeaths[index, 1] = labels_other_key[f + 1, 2]
        wb_wt_CheckDeaths[index, 2] = sexlabel[s]
        wb_wt_CheckDeaths[index, 3] = pclabel[a]}}}
  
  # Construct Dataframe
  wb_wt_Log = data.frame(wb_wt_Log)
  wb_wt_CheckMig = data.frame(wb_wt_CheckMig)
  wb_wt_CheckDeaths = data.frame(wb_wt_CheckDeaths)
  colnames(wb_wt_Log) = c("Projection interval", "Population-at-risk iterations", "Mig adjustment iterations")
  colnames(wb_wt_CheckMig) = c("Year-Cohort", "Sex", "Period-Cohort", Areaname, Areaname, "National net mig")
  colnames(wb_wt_CheckDeaths) = c("Year-Cohort", "Sex", "Period-Cohort", Areaname, "National deaths")
  
  # Return dataframe
  return (list("Log" = wb_wt_Log, "Mig" = wb_wt_CheckMig, "Deaths" = wb_wt_CheckDeaths))
}

# Function for setting initial value of jump-off year population
readIniPop <- function(final, numareas, numages, ERP){

  # Create empty array for storing population data
  Population = array(rep(0, (final + 1) * numareas * 2 * numages),
                     dim = c((final + 1), numareas, 2, numages))

  # record initial (jump-off year) population from ERP dataset
  for (i in 1:numareas){
    for (s in 1:2){
      for (pc in 1:numages){
        Population[1, i, s, pc] = ERP[1, i, s, pc]}}}

  # Return initial population dataset (the rest 3 levels are used for storing the projection)
  return (Population)
}

# Function for recording the total population in jump-off year in each area
readIniTPop <- function(final, numareas, numages, Population){
  
  # Create empty array for collecting total population
  totPopulation = array(rep(0, (final + 1) * numareas), dim = c((final + 1), numareas))
  
  # Record jump-off year total population in each area
  for (i in 1:numareas){
    totPopulation[1, i] = 0
    for (s in 1:2){
      for (a in 1:numages){
        totPopulation[1, i] = totPopulation[1, i] + Population[1, i, s, a]}}}
  
  # Return total population of jump-off year (the rest 3 levels are used for storing the projection)
  return (totPopulation)
}

# Function for generating scaled inward and outward migration projection'''
NetMigAdjustment2 <- function(labely, prelimIMpc, prelimOMpc, NatN, requiredN, LocalPop0, LocalDpc,
                              scaledIM, scaledOM, numareas, lastage, maxziter, smallnumber, wb_wt_Log){
  totN1 = 0
  for (i in 1:numareas){
    totN1 = totN1 + requiredN[i]}

  totN2 = 0
  for (s in 1:2){
    for (pc in 1:(lastage + 1)){
      totN2 = totN2 + NatN[s, pc]}}
  
  IM1 = array(rep(0, numareas * 2 * (lastage + 1)), dim = c(numareas, 2, (lastage + 1)))
  OM1 = array(rep(0, numareas * 2 * (lastage + 1)), dim = c(numareas, 2, (lastage + 1)))
  for (i in 1:numareas){
    for (s in 1:2){
      for (pc in 1:(lastage + 1)){
        IM1[i, s, pc] = prelimIMpc[i, s, pc]
        OM1[i, s, pc] = prelimOMpc[i, s, pc]}}}

  for (z in 1:maxziter){
    
    # Record the number of iterations required
    wb_wt_Log[labely - 1, 3] = z
  
    # Calculate inward & outward migration by sex and period-cohort summed over 
    totIM1 = array(rep(0, 2 * (lastage + 1)), dim = c(2, (lastage + 1)))
    totOM1 = array(rep(0, 2 * (lastage + 1)), dim = c(2, (lastage + 1)))
    for (i in 1:numareas){
      for (s in 1:2){
        for (pc in 1:(lastage + 1)){
          totIM1[s, pc] = totIM1[s, pc] + IM1[i, s, pc]
          totOM1[s, pc] = totOM1[s, pc] + OM1[i, s, pc]}}}
  
    # Calculate the scaling factors
    sf = array(rep(0, 2 * (lastage + 1)), dim = c(2, (lastage + 1)))
    for (s in 1:2){
      for (pc in 1:(lastage + 1)){
        if (totIM1[s, pc] > 0){
          sf[s, pc] = (NatN[s, pc] + (((NatN[s, pc] ^ 2) + (4 * totIM1[s, pc] * totOM1[s, pc])) ^ 0.5)) / (2 * totIM1[s, pc])}}}
  
    # Make inward & outward migration flows consistent with national net migration by sex-age cohorts
    IM2 = array(rep(0, numareas * 2 * (lastage + 1)), dim = c(numareas, 2, (lastage + 1)))
    OM2 = array(rep(0, numareas * 2 * (lastage + 1)), dim = c(numareas, 2, (lastage + 1)))
    for (i in 1:numareas){
      for (s in 1:2){
        for (pc in 1:(lastage + 1)){
          IM2[i, s, pc] = IM1[i, s, pc] * sf[s, pc]
          OM2[i, s, pc] = OM1[i, s, pc] / sf[s, pc]}}}
  
    # Calculate inward & outward migration for each area summed over sex and age cohort
    totIM2 = rep(0, numareas)
    totOM2 = rep(0, numareas)
    for (i in 1:numareas){
      for (s in 1:2){
        for (pc in 1:(lastage + 1)){
          totIM2[i] = totIM2[i] + IM2[i, s, pc]
          totOM2[i] = totOM2[i] + OM2[i, s, pc]}}}
  
    # Calculate required total inward migration
    totin = rep(0, numareas)
    for (i in 1:numareas){
      totin[i] = requiredN[i] + totOM2[i]}
  
    # Adjust inward migration so that it gives the required local net migration total
    IM3 = array(rep(0, numareas * 2 * (lastage + 1)), dim = c(numareas, 2, (lastage + 1)))
    OM3 = array(rep(0, numareas * 2 * (lastage + 1)), dim = c(numareas, 2, (lastage + 1)))
    for (i in 1:numareas){
      for (s in 1:2){
        for (pc in 1:(lastage + 1)){
          IM3[i, s, pc] = IM2[i, s, pc] * totin[i] / totIM2[i]
          OM3[i, s, pc] = OM2[i, s, pc]}}}
  
    # Check to ensure that net migration does not given a -ve population and adjust migration flows if necessary
    IM4 = array(rep(0, numareas * 2 * (lastage + 1)), dim = c(numareas, 2, (lastage + 1)))
    OM4 = array(rep(0, numareas * 2 * (lastage + 1)), dim = c(numareas, 2, (lastage + 1)))
    for (i in 1:numareas){
      for (s in 1:2){
        for (pc in 1:(lastage + 1)){
          tempcheck = LocalPop0[i, s, pc] - LocalDpc[i, s, pc] - OM3[i, s, pc] + IM3[i, s, pc]
          if (tempcheck < 0){
            OM4[i, s, pc] = OM3[i, s, pc] + 0.6 * tempcheck
            IM4[i, s, pc] = IM3[i, s, pc] - 0.6 * tempcheck
            print("Warning: Migration flows had to be adjusted to avoid negative population")}
          else{
            IM4[i, s, pc] = IM3[i, s, pc]
            OM4[i, s, pc] = OM3[i, s, pc]}}}}
  
    # Check the difference between IM, OM 1 and 4 matrices
    diffr = array(rep(0, 2 * numareas * 2 * (lastage + 1)), dim = c(2, numareas, 2, (lastage + 1)))
    for (i in 1:numareas){
      for (s in 1:2){
        for (pc in 1:(lastage + 1)){
          diffr[1, i, s, pc] = abs(IM4[i, s, pc] - IM1[i, s, pc])
          diffr[2, i, s, pc] = abs(OM4[i, s, pc] - OM1[i, s, pc])}}}
  
    checkOK = array(rep(0, 2 * numareas * 2 * (lastage + 1)), dim = c(2, numareas, 2, (lastage + 1)))
    totcheckOK = 0
    for (k in 1:2){
      for (i in 1:numareas){
        for (s in 1:2){
          for (pc in 1:(lastage + 1)){
            if (diffr[k, i, s, pc] < smallnumber){
              checkOK[k, i, s, pc] = 0}
            else{
              checkOK[k, i, s, pc] = 1
              totcheckOK = totcheckOK + checkOK[k, i, s, pc]}}}}}
  
    # Check convergence
    if (totcheckOK == 0){
      for (i in 1:numareas){
        for (s in 1:2){
          for (pc in 1:(lastage + 1)){
            scaledIM[i, s, pc] = IM4[i, s, pc]
            scaledOM[i, s, pc] = OM4[i, s, pc]}}}
      return (list("IM" = scaledIM, "OM" = scaledOM, "Log" = wb_wt_Log))}
    else{
      IM1 = array(rep(0, numareas * 2 * (lastage + 1)), dim = c(numareas, 2, (lastage + 1)))
      OM1 = array(rep(0, numareas * 2 * (lastage + 1)), dim = c(numareas, 2, (lastage + 1)))
      for (i in 1:numareas){
        for (s in 1:2){
          for (pc in 1:(lastage + 1)){
            IM1[i, s, pc] = IM4[i, s, pc]
            OM1[i, s, pc] = OM4[i, s, pc]}}}}
    # Continue looping
  }
}