import numpy

'''Function for generating 3-Dimension jumpoffERP Array'''
def jumpoff(numareas, numages, jumpoffERP, sheet_agesex, set_year_female, set_year_male):
    row = 3
    for i in range(numareas):
        for a in range(numages):
            row += 1
            jumpoffERP[i, 0, a] = sheet_agesex.cell_value(row, set_year_female)
            jumpoffERP[i, 1, a] = sheet_agesex.cell_value(row, set_year_male)

    # Return jumoffERP for further calculation
    return jumpoffERP

'''Function for generating 1-Dimension xTFR Array'''
def xTFR(jumpoffERP, result_xTFR, numareas):

    for i in range(numareas):
        # Related age_sex group
        male04 = jumpoffERP[i, 1, 0]
        female04 = jumpoffERP[i, 0, 0]
        female_1519 = jumpoffERP[i, 0, 3]
        female_2024 = jumpoffERP[i, 0, 4]
        female_2529 = jumpoffERP[i, 0, 5]
        female_3034 = jumpoffERP[i, 0, 6]
        female_3539 = jumpoffERP[i, 0, 7]
        female_4044 = jumpoffERP[i, 0, 8]
        female_4549 = jumpoffERP[i, 0, 9]
        
        # Subjective definition for simplifying formula
        new_born = male04 + female04
        most_fert = female_2529 + female_3034
        all_fert = female_1519 + female_2024 + female_2529 + female_3034 + female_3539 + female_4044 + female_4549

        # Calculate result for xTFR in each area
        result_xTFR[i] = (10.65 - (12.55 * most_fert / all_fert)) * new_born / all_fert

    # Return xTRF result for storing into Fertility Sheet
    return result_xTFR

#################################################################################################################################################################
'''Function to clear the SmallAreaInputs before running the PrepareData Function if Input Data Already Exists'''
def clear_Input(worksheet):

    # Clear All row except of the Title Row
    worksheet.delete_rows(2, worksheet.max_row)

    # return to main function
    return

'''Function for collecting Area Codes, Names, Age Groups and Perid-Cohort Labels'''
def readCNAPC(numareas, numages, lastage, sheet_label):

    # Create empty Area Code, Name, Age Group List for collection
    AreaCode = [None] * numareas
    AreaName = [None] * numareas
    AgeLabel = [None] * numages
    PCLabel = [None] * (lastage + 1)

    # Collect Area Code and Name
    row = 3
    for i in range(numareas):
        row += 1
        AreaCode[i] = int(sheet_label.cell_value(row, 1))
        AreaName[i] = sheet_label.cell_value(row, 2)
    
    # Collect Age Group Label
    age_row = 3
    for i in range(numages):
        age_row += 1
        AgeLabel[i] = sheet_label.cell_value(age_row, 5)
    
    # Collect Period-Cohort Label
    pc_row = 3
    for i in range(lastage+1):
        pc_row += 1
        PCLabel[i] = sheet_label.cell_value(pc_row, 8)

    # Return AreaCode, AreaName, AgeLabel
    return AreaCode, AreaName, AgeLabel, PCLabel

'''Function for collecting Sex, Year, Interval Labels'''
def readSYI(final, sheet_label):

    # Create empty Sex, Year, Interval Label List for collection
    SexLabel = ['Females', 'Males']
    YearLabel = [None] * (final + 2)
    IntervalLabel = [None] * (final + 1)

    # Collect Year Labels
    year_col = 2
    for i in range(final + 2):
        year_col += 1
        YearLabel[i] = int(sheet_label.cell_value(2259, year_col))
    
    # Collect Interval Labels
    interval_col = 2
    for i in range(final + 1):
        interval_col += 1
        IntervalLabel[i] = sheet_label.cell_value(2262, interval_col)
    
    # Return SexLabel, YearLabel, IntervalLabel
    return (SexLabel, YearLabel, IntervalLabel)

'''Function for collecting Estimated & Projected Small Area Total Population'''
def readSATP(intervals, sheet_SmallAreaTotals, numareas):

    # Create array for storing Estimated & Projected Small Area Total Population Results
    smallAreaTotalPop = numpy.zeros((intervals, numareas))
    smallAreaTotalPop[:] = numpy.nan
    
    # Set start column for inserting / reading data
    row = 3
    if (intervals == 2):
        int_col = 2
    else:
        int_col = 3
    
    # Collect Small Area Total Population
    for i in range(numareas):
        row += 1
        SATP_col = int_col
        for a in range(intervals):
            SATP_col += 1
            smallAreaTotalPop[a, i] = sheet_SmallAreaTotals.cell_value(row, SATP_col)
    
    # Return Small Area Total Population for specific range of years
    return smallAreaTotalPop

'''Function for recording ERP (Estimated Resident Population) in small area with age & sex information'''
def readERP(intervals, numareas, numages, sheet_agesex, project):

    # Create empty array for storing ERP in small area by age and sex
    infoERP = numpy.zeros((intervals, numareas, 2, numages))
    infoERP[:] = numpy.nan

    # Collect data
    for year in range(intervals):
        row = 3
        for i in range(numareas):
            for a in range(numages):
                row += 1
                if (project):
                    index = year + 1
                else:
                    index = year
                # Females
                infoERP[year, i, 0, a] = sheet_agesex.cell_value(row, 6 + index)
                # Males
                infoERP[year, i, 1, a] = sheet_agesex.cell_value(row, 4 + index)
    
    # Return ERP data with given year range(s)
    return infoERP

'''Function to record TFR information'''
def readTFR(final, numareas, sheet_fertility):

    # Create empty array for stroing TFR information
    infoTFR = numpy.zeros((final+1, numareas))
    infoTFR[:] = numpy.nan

    # Collect TFR values
    row = 5
    col = 3
    for i in range(numareas):
        row += 1
        for a in range(final+1):
            infoTFR[a, i] = sheet_fertility.cell_value(row, col)
    
    # Return TFR array
    return infoTFR

'''Function to read in data for ASFRs and Preliminary ASFRs (Area specific Fertility Rate)'''
def readASFR(numareas, sheet_fertility, age_groups):

    # Create empty list and array for storing data
    ASFR = [None] * age_groups
    prelimASFR = numpy.zeros((numareas, age_groups))
    prelimASFR[:] = numpy.nan

    # Collect data for ASFRs
    row = 2258
    for i in range(age_groups):
        row += 1
        ASFR[i] = sheet_fertility.cell_value(row, 1)
    
    # Collect data for preliminary ASFRs
    row = 5
    for i in range(numareas):
        row += 1
        if sheet_fertility.cell_value(row, 16) != '':
            col = 15
            for a in range(age_groups):
                col += 1
                prelimASFR[i, a] = sheet_fertility.cell_value(row, col)
        else:
            col = 15
            for a in range(age_groups):
                col += 1
                prelimASFR[i, a] = 0
    
    # Return ASFRs and Preliminary ASFRs data
    return ASFR, prelimASFR

'''Function to collect Life Expectancy at Birth Assumptions, Mortality Surface Information'''
def readLEMS(final, numareas, lastage, nLxMS, sheet_mortality):
    
    # Create empty array for storing Life expectancy and Mortality Surface
    eO = numpy.zeros((final + 1, numareas, 2))
    eO[:] = numpy.nan
    MS_nLx = numpy.zeros((nLxMS, 2, lastage + 1))
    MS_nLx[:] = numpy.nan
    MS_TO = numpy.zeros((nLxMS, 2))
    MS_TO[:] = numpy.nan

    # Collect Life Expectancy Data
    row_eO = 5
    for i in range(numareas):
        col_eO = 2
        for a in range(final + 1):
            col_eO += 1
            eO[a, i, 0] = sheet_mortality.cell_value(row_eO, col_eO)
            eO[a, i, 1] = sheet_mortality.cell_value(row_eO, col_eO + 13)
    
    # Collect Mortality Surface Data
    row_MS = 2259
    for i in range(lastage + 1):
        row_MS += 1
        col_MS = 4
        for a in range(nLxMS):
            col_MS += 1
            MS_nLx[a, 0, i] = sheet_mortality.cell_value(row_MS, col_MS)
            MS_nLx[a, 1, i] = sheet_mortality.cell_value(row_MS + 21, col_MS)
    
    col_TO = 4
    for i in range(nLxMS):
        col_TO += 1
        MS_TO[i, 0] = sheet_mortality.cell_value(2278, col_TO)
        MS_TO[i, 1] = sheet_mortality.cell_value(2299, col_TO)
    
    # Return Life Expectancy and Mortality Surface Data
    return eO, MS_nLx, MS_TO

'''Function of data collection for ASMRs (Area specific Mortality Rate) and Migration Information'''
def readMASMR(numareas, lastage, sheet_migration):
    
    # Create empty list / array for storing data of migration and for ASMRs
    totmig = [None] * numareas
    modelASMR = numpy.zeros((2, lastage + 1))
    modelASMR[:] = numpy.nan

    # Collect Migration Data
    row_mig = 4
    for i in range(numareas):
        row_mig += 1
        # Default Value
        totmig[i] = sheet_migration.cell_value(row_mig, 3)
        # User-supplied Value
        if sheet_migration.cell_value(row_mig, 4) != '':
            totmig[i] = sheet_migration.cell_value(row_mig, 4)
    
    # Collect ASMRs Data
    row_ASMR = 10
    for i in range(lastage + 1):
        row_ASMR += 1
        col = 6
        for a in range(2):
            col += 1
            modelASMR[a, i] = sheet_migration.cell_value(row_ASMR, col)
    
    # Return Data of Migration and ASMRs
    return totmig, modelASMR

#################################################################################################################################################################
'''Function for Generating Input Data for ASFR (Area specific Fertility Rate)'''
def inputASFR(final, numareas, age_groups, sheet_fertility, prelimASFR, modelASFR, TFR, modelTFR):
    ASFR = numpy.zeros((final + 1, numareas, age_groups))
    
    # Collect ASFR input data
    for y in range(final + 1):
        row = 5
        for i in range(numareas):
            row += 1
            if sheet_fertility.cell_value(row, 16) != "":
                prelimtot = 0
                for a in range(age_groups):
                    prelimtot += prelimASFR[i, a]
                for a in range(age_groups):
                    ASFR[y, i, a] = prelimASFR[i, a] * TFR[y, i] / (prelimtot * 5)
            else:
                for a in range(age_groups):
                    ASFR[y, i, a] = modelASFR[a] * TFR[y, i] / modelTFR
    
    # Return ASFR input Data
    return ASFR

'''Function for estimating base periold birth by sex'''
def inputbirth(numareas, age_groups, ASFR_data, ERP, SRB):

    tempbirths = [None] * age_groups            # Amount of Each Age group female could give birth to
    tempbirthssex = numpy.zeros((numareas, 2))  # Total estimate birth in different sex
    temptotbirths = [None] * numareas           # Total estimate birth

    # Collect estimated total birth population and sex birth population
    for i in range(numareas):
        temptotbirths[i] = 0

        for a in range(age_groups):
            tempbirths[a] = ASFR_data[0, i, a] * 2.5 * (ERP[0, i, 0, a + 3] + ERP[1, i, 0, a + 3])
            temptotbirths[i] += tempbirths[a]

        tempbirthssex[i, 0] = temptotbirths[i] * (100 / (SRB + 100))
        tempbirthssex[i, 1] = temptotbirths[i] * (SRB / (SRB + 100))
    
    # Return tempbirthssex, temptotbirths estimation data
    return tempbirthssex, temptotbirths

'''Function for generating ASDR Input Data for further usage'''
def inputASDR(final, numareas, lastage, nLxMS, eO, MS_TO, MS_nLx):
    
    # Create empty array for storing ASDR input data
    ASDR = numpy.zeros((final + 1, numareas, 2, lastage + 1))

    for i in range(numareas):
        Lower_TO = numpy.zeros((final + 1, 2))
        Upper_TO = numpy.zeros((final + 1, 2))
        Lower_num = numpy.zeros((final + 1, 2))
        Upper_num = numpy.zeros((final + 1, 2))
        proportion = numpy.zeros((final + 1, 2))

        # Find out where in the mortality surface each eO lies
        for y in range(final + 1):
            for s in range(2):
                for z in range(nLxMS - 1):
                    e0_100k = eO[y, i, s] * 100000
                    if (e0_100k >= MS_TO[z, s] and e0_100k < MS_TO[z + 1, s]):
                        Upper_TO[y, s] = MS_TO[z + 1, s]
                        Upper_num[y, s] = z + 1
                        Lower_TO[y, s] = MS_TO[z, s]
                        Lower_num[y, s] = z
                        proportion[y, s] = (e0_100k - Lower_TO[y, s]) / (Upper_TO[y, s] - Lower_TO[y, s])
                        break

        # Create nLx values for each eO
        nLx = numpy.zeros((final + 1, 2, lastage + 1))
        for y in range(final + 1):
            for s in range(2):
                lower = 0 # Ensure there is lower value without the above loop
                upper = 0 # Ensure there is upper value without the above loop
                lower = int(Lower_num[y, s])
                upper = int(Upper_num[y, s])
                for a in range(lastage + 1):
                    MS_low = MS_nLx[lower, s, a]
                    MS_up = MS_nLx[upper, s, a]
                    nLx[y, s, a] = MS_low + proportion[y, s] * (MS_up - MS_low)
        
        # Create Period-Cohort ASDRs (Area Specific Death Rates)
        for y in range(final + 1):
            for s in range(2):
                
                # Record Area Specific Death Rate for age groups excluding 0-4 and 85+
                for pc in range(1, lastage):
                    younger = nLx[y, s, pc - 1]
                    elder = nLx[y, s, pc]
                    ASDR[y, i, s, pc] = (younger - elder) / (5 / 2 * (younger + elder))
                
                # Record ASDR for 0-4
                ASDR[y, i, s, 0] = ((5 * 100000) - nLx[y, s, 0]) / (5 / 2 * nLx[y, s, 0])
                # Record ASDR for 85+
                younger = nLx[y, s, lastage - 1]
                elder = nLx[y, s, lastage]
                ASDR[y, i, s, lastage] = ((younger + elder) - elder) / (5 / 2 * (younger + elder + elder))
    
    # Return Area Specific Death Rates Input Data
    return ASDR

'''Function for calculating the estimated bas period deaths for the population accounts'''
def inputdeath(numareas, lastage, ASDR, ERP):

    tempdeaths = numpy.zeros((numareas, 2, lastage + 1))  # Number of death in each age-sex group
    temptotdeaths = [None] * numareas                     # Total estimate death

    # Record number of death in each age-sex cohort
    for i in range(numareas):
        for s in range(2):
            # Record estimated Deaths for age-sex groups excluding 0-4 and 85+
            for pc in range(1, lastage):
                tempdeaths[i, s, pc] = ASDR[0, i, s, pc] * 2.5 * (ERP[0, i, s, pc - 1] + ERP[1, i, s, pc])
            # Record Deaths for 0-4
            tempdeaths[i, s, 0] = ASDR[0, i, s, 0] * 2.5 * ERP[1, i, s, 0]
            # Record Deaths for 85+
            younger = ERP[0, i, s, lastage - 1]
            elder_f = ERP[0, i, s, lastage]
            elder_m = ERP[1, i, s, lastage]
            tempdeaths[i, s, lastage] = ASDR[0, i, s, lastage] * 2.5 * (younger + elder_f + elder_m)

        # Record estimated total death
        temptotdeaths[i] = 0
        for s in range(2):
            for pc in range(lastage + 1):
                temptotdeaths[i] += tempdeaths[i, s, pc]
        
    # Return tempdeaths, temptotdeaths estimation data
    return tempdeaths, temptotdeaths

'''Function for calculating migration input data'''
def inputMigration(final, numareas, lastage, modelASMR, ERP, totmig, TotPop, temptotdeaths, temptotbirths, tempdeaths, tempbirthssex, ASDR, Areaname, sexlabel, pclabel, intervallabel):

    # Create empty array for storing migration related data
    ASOMR = numpy.zeros((final + 1, numareas, 2, lastage + 1))  # Area Specific Out Migration Ratio
    inward = numpy.zeros((final + 1, numareas, 2, lastage + 1)) # Inward migration amount
    outward = numpy.zeros((numareas, 2, lastage + 1))           # Outward migration amount
    prelimmig = numpy.zeros((numareas, 2, lastage + 1))         # Prelimium migration
    scaledmig = numpy.zeros((numareas, 2, lastage + 1))         # Scaled migration
    prelimbaseinward = numpy.zeros((numareas, 2, lastage + 1))  # Prelimium inward migration
    prelimbaseoutward = numpy.zeros((numareas, 2, lastage + 1)) # Prelimium outward migration
    cohortnetmig = numpy.zeros((numareas, 2, lastage + 1))      # Cohort net migration
    adjnetmig = numpy.zeros((numareas, 2, lastage + 1))         # Adjuested net migration
    scalingfactor2a = numpy.zeros((numareas, 2, lastage + 1))   # Raw scaling factor
    scalingfactor2b = numpy.zeros((numareas, 2, lastage + 1))   # Smoothed scaling factor

    # Collecting Migration Input Data
    for i in range(numareas):

        # Generate preliminary migration turnover by sex-age cohort based on model rates
        for s in range(2):
            # Generate preliminary migration turnover for all sex-age cohorts exclude 0-4, 85+
            for pc in range(1, lastage):
                prelimmig[i, s, pc] = modelASMR[s, pc] * 2.5 * (ERP[0, i, s, pc - 1] + ERP[1, i, s, pc])

            # Generate preliminary migration turnover for 0-4
            prelimmig[i, s, 0] = modelASMR[s, 0] * 2.5 * ERP[1, i, s, 0]

            # Generate preliminary migration turnover for 85+
            prelimmig[i, s, lastage] = modelASMR[s, lastage] * 2.5 * (ERP[0, i, s, lastage - 1] + ERP[0, i, s, lastage] + ERP[1, i, s, lastage])

        # Generate total preliminarymigration turnover value
        totprelimmig = 0
        for s in range(2):
            for pc in range(lastage + 1):
                totprelimmig += prelimmig[i, s, pc]
        
        # Collecting scaled preliminary migration by sex-age cohort to migration turnover 
        # as calculated by the crude migration turnover rate; and then divide by 2 to give the initial inward and outward migration
        for s in range(2):
            for pc in range(lastage + 1):
                scaledmig[i, s, pc] = prelimmig[i, s, pc] * (0.5 * totmig[i]) / totprelimmig

        # Calculate residual net migration for base period in each area i
        basenetmig = TotPop[1, i] - TotPop[0, i] + temptotdeaths[i] - temptotbirths[i]

        # Calculate the scaling factor to estimate separate inward and outward migration totals which are consistent with residual net migration
        scalingfactor1 = (basenetmig + (basenetmig ** 2 + (4 * 0.5 * totmig[i] * 0.5 * totmig[i])) ** 0.5) / (2 * 0.5 * totmig[i])

        # Calculate preliminary inward and outward migration by sex-age cohort (Different with the Original Value in 'Account')
        for s in range(2):
            for pc in range(lastage + 1):
                prelimbaseinward[i, s, pc] = scaledmig[i, s, pc] * scalingfactor1
                prelimbaseoutward[i, s, pc] = scaledmig[i, s, pc] / scalingfactor1

        # Calculate cohort-specific residual net migration (for adjustment usage)
        for s in range(2):
            # Generate cohort-specific residual net migration for all sex-age cohorts exclude 0-4, 85+
            for pc in range(1, lastage):
                cohortnetmig[i, s, pc] = ERP[1, i, s, pc] - ERP[0, i, s, pc - 1] + tempdeaths[i, s, pc]
            
            # Generate cohort-specific residual net migration for 0-4
            cohortnetmig[i, s, 0] = ERP[1, i, s, 0] - tempbirthssex[i, s] + tempdeaths[i, s, 0]

            # Generate cohort-specific residual net migration for 85+
            cohortnetmig[i, s, lastage] = ERP[1, i, s, lastage] - (ERP[0, i, s, lastage - 1] + ERP[0, i, s, lastage]) + tempdeaths[i, s, lastage]

        # Adjusted cohort-specific net migration with averaged over sex for selected ages
        for s in range(2):
            for pc in range(14):
                adjnetmig[i, s, pc] = 0.5 * (cohortnetmig[i, 0, pc] + cohortnetmig[i, 1, pc])
            for pc in range(14, lastage + 1):
                adjnetmig[i, s, pc] = cohortnetmig[i, s, pc]
        
        # Calculate scaling factors to adjust directional migration (Raw & Smoothed)
        # (a) Raw Scaling factor
        for s in range(2):
            for pc in range(lastage + 1):
                if prelimbaseinward[i, s, pc] > 0:
                    adjmig = adjnetmig[i, s, pc]
                    preinward = prelimbaseinward[i, s, pc]
                    preoutward = prelimbaseoutward[i, s, pc]
                    scalingfactor2a[i, s, pc] = (adjmig + (adjmig ** 2 + (4 * preinward * preoutward)) ** 0.5) / (2 * preinward)
                else:
                    scalingfactor2a[i, s, pc] = 1
        
        # (b) Smoothed Scaling factor
        for s in range(2):
            for pc in range(lastage + 1):
                if (pc < 10):
                    scalingfactor2b[i, s, pc] = scalingfactor2a[i, s, pc]
                elif (pc == lastage):
                    scalingfactor2b[i, s, pc] = 1 / 2 * (scalingfactor2a[i, s, pc - 1] + scalingfactor2a[i, s, pc])
                else:
                    scalingfactor2b[i, s, pc] = 1 / 3 * (scalingfactor2a[i, s, pc - 1] + scalingfactor2a[i, s, pc] + scalingfactor2a[i, s, pc + 1])
                
                # Further guarantee of no too extreme scaling factor value
                if (scalingfactor2b[i, s, pc] > 10):
                    scalingfactor2b[i, s, pc] = 10
                elif(scalingfactor2b[i, s, pc] < 0.1):
                    scalingfactor2b[i, s, pc] = 0.1

        # Adjust inward and outward migration with smoothed scaling factor
        for s in range(2):
            for pc in range(lastage + 1):
                inward[0, i, s, pc] = prelimbaseinward[i, s, pc] * scalingfactor2b[i, s, pc]
                outward[i, s, pc] = prelimbaseoutward[i, s, pc] / scalingfactor2b[i, s, pc]
        
        # Smoothing inward migration age profiles at selected age groups
        intemp = numpy.zeros((2, lastage + 1))
        for s in range(2):
            for pc in range(lastage + 1):
                intemp[s, pc] = inward[0, i, s, pc]
        for s in range(2):
            for pc in range(0, lastage - 1):
                inward[0, i, s, pc] = intemp[s, pc]
            for pc in range(lastage - 1, lastage + 1):
                inward[0, i, s, pc] = 1 / 2 * (intemp[0, pc] + intemp[1, pc])
        
        # Calculate ASOMR (Area Specific Out Migration Rate)
        for s in range(2):
            # Record ASOMR input for all sex-age cohort exclude 0-4 & 85+
            for pc in range(1, lastage):
                if (ERP[0, i, s, pc - 1] + ERP[1, i, s, pc] > 0):
                    ASOMR[0, i, s, pc] = outward[i, s, pc] / (2.5 * (ERP[0, i, s, pc - 1] + ERP[1, i, s, pc]))
                else:
                    ASOMR[0, i, s, pc] = 0
                    print("Warning: Base period out-migration rate canoot be calculated due to zero population")
                    print("Aera = ", Areaname[i], " & Sex = ", sexlabel[s], " & Period-Cohort = ", pclabel[pc])
            
            # Record ASOMR input for sex-age cohort 0-4
            if (ERP[1, i, s, 0] > 0):
                ASOMR[0, i, s, 0] = outward[i, s, 0] / (2.5 * ERP[1, i, s, 0])
            else:
                ASOMR[0, i, s, 0] = 0
                print("Warning: Base period out-migration rate canoot be calculated due to zero population")
                print("Aera = ", Areaname[i], " & Sex = ", sexlabel[s], " & Period-Cohort = ", pclabel[0])
                
            # Record ASOMR inpurt for sex age ohort 85+
            if (ERP[0, i, s, lastage - 1] + ERP[0, i, s, lastage]) > 0:
                ASOMR[0, i, s, lastage] = outward[i, s, lastage] / (2.5 * (ERP[0, i, s, lastage - 1] + ERP[0, i, s, lastage] + ERP[1, i, s, lastage]))
            else:
                ASOMR[0, i, s, lastage] = 0
        
        # Smoothing ASOMR at selected age groups
        outtemp = numpy.zeros((2, lastage + 1))
        for s in range(2):
            for pc in range(lastage + 1):
                outtemp[s, pc] = ASOMR[0, i, s, pc]
        # Start smoothing
        for s in range(2):
            # When for the first 11 groups, keep ASOMR unchanged
            for pc in range(11):
                ASOMR[0, i, s, pc] = outtemp[s, pc]
            
            # Smooth the rest of the elder age groups ASOMR
            for pc in range(11, lastage + 1):
                ASOMR[0, i, s, pc] = 1 / 2 * (outtemp[0, pc] + outtemp[1, pc])
        
        # Set default ASOMR and Inward migration projection on the future 3 year intervals (20-25, 25-30, 40-35)
        for y in range(1, final + 1):
            for s in range(2):
                for pc in range(lastage + 1):
                    ASOMR[y, i, s, pc] = ASOMR[0, i, s, pc]
                    inward[y, i, s, pc] = inward[0, i, s, pc]
        
        # Check that ASOMRs + ASDR do not exceed 0.4 & Adjust ASOMRs if necessary
        # Helps to avoid negative population if there is no inward migration because pop-at-risk = 2.5 * (Pop0 + Pop1)
        for y in range(final + 1):
            for s in range(2):
                for pc in range(lastage + 1):
                    if (ASOMR[y, i, s, pc] + ASDR[y, i, s, pc] > 0.5):
                        temprate = ASOMR[y, i, s, pc]
                        ASOMR[y, i, s, pc] = 0.4 - ASDR[y, i, s, pc]
                        print("Warning: Excessive outward migration rate reducing")
                        print("Avoid having outward migration rate + age-specific death rate from exceeding 0.4")
                        print("Old rate = ", temprate, " & New rate = ", ASOMR[y, i, s, pc], " & Area = ", Areaname[i], " & Sex = ", sexlabel[s], " & Period-Cohort = ", pclabel[pc], " & Interval = ", intervallabel[y])
    
    # Return migration input data
    return ASOMR, inward, outward, prelimmig, scaledmig, prelimbaseinward, prelimbaseoutward, cohortnetmig, adjnetmig, scalingfactor2a, scalingfactor2b

#################################################################################################################################################################
'''Function for writing inputs Accounts'''
def writeAccount(wb_wt_Accounts, numareas, Areaname, lastage, sexlabel, pclabel, tempbirthssex, inward, outward, prelimmig, scaledmig, prelimbaseinward, prelimbaseoutward, cohortnetmig, adjnetmig, scalingfactor2a, scalingfactor2b, ERP, tempdeaths):
    
    # Record headings of Accounts
    wb_wt_Accounts.cell(3, 3).value = "ERP(t-5)"
    wb_wt_Accounts.cell(3, 4).value = "Deaths"
    wb_wt_Accounts.cell(3, 5).value = "Net mig"
    wb_wt_Accounts.cell(3, 6).value = "ERP(t)"
    wb_wt_Accounts.cell(3, 8).value = "Prelim mig (model rates x pop-at-risk)"
    wb_wt_Accounts.cell(3, 9).value = "Scaled mig (adjusted for pop size)"
    wb_wt_Accounts.cell(3, 10).value = "Prelim in (consistent with total net mig)"
    wb_wt_Accounts.cell(3, 11).value = "Prelim out (consistent with total net mig)"
    wb_wt_Accounts.cell(3, 12).value = "scaling factor a (unsmoothed)"
    wb_wt_Accounts.cell(3, 13).value = "scaling factor b (smoothed)"
    wb_wt_Accounts.cell(3, 14).value = "In-mig (consistent with cohort net mig)"
    wb_wt_Accounts.cell(3, 15).value = "Out-mig (consistent with cohort net mig)"
    wb_wt_Accounts.cell(3, 16).value = "Adjusted net mig"

    # Record input data of Accounts
    row = 3
    for i in range(numareas):

        # Record area name for each sector of input data in Accounts Sheet
        row += 1
        wb_wt_Accounts.cell(row, 1).value = str(i + 1) + " " + Areaname[i]

        for s in range(2):
            for pc in range(lastage + 1):
                row += 1

                # Record sex and age-cohort label
                wb_wt_Accounts.cell(row, 1).value = sexlabel[s]
                wb_wt_Accounts.cell(row, 2).value = pclabel[pc]

                # Record ERP(t-5)
                if (pc == 0):
                    wb_wt_Accounts.cell(row, 3).value = tempbirthssex[i, s]
                elif (pc > 0 and pc < lastage):
                    wb_wt_Accounts.cell(row, 3).value = ERP[0, i, s, pc]
                else:
                    wb_wt_Accounts.cell(row, 3).value = ERP[0, i, s, lastage - 1] + ERP[0, i, s, lastage]
                
                wb_wt_Accounts.cell(row, 4).value = tempdeaths[i, s, pc]         # Record estimated death
                wb_wt_Accounts.cell(row, 5).value = cohortnetmig[i, s, pc]       # Record cohort net migration
                wb_wt_Accounts.cell(row, 6).value = ERP[1, i, s, pc]             # Record estimated resident population
                wb_wt_Accounts.cell(row, 8).value = prelimmig[i, s, pc]          # Record preliminary migration
                wb_wt_Accounts.cell(row, 9).value = scaledmig[i, s, pc]          # Record scaled migration
                wb_wt_Accounts.cell(row, 10).value = prelimbaseinward[i, s, pc]  # Record preliminary inward migration
                wb_wt_Accounts.cell(row, 11).value = prelimbaseoutward[i, s, pc] # Record preliminary outward migration
                wb_wt_Accounts.cell(row, 12).value = scalingfactor2a[i, s, pc]   # Record raw scaling factor 
                wb_wt_Accounts.cell(row, 13).value = scalingfactor2b[i, s, pc]   # Record smoothed scaling factor
                wb_wt_Accounts.cell(row, 14).value = inward[0, i, s, pc]         # Record inward migration
                wb_wt_Accounts.cell(row, 15).value = outward[i, s, pc]           # Record outward migration
                wb_wt_Accounts.cell(row, 16).value = adjnetmig[i, s, pc]         # Record new migration
        
        # Create an empty row for spliting data table
        row += 1
    
    # Return the Accounts Sheet
    return wb_wt_Accounts

'''Sub-function for writing input data (Headings) to SmallAreaInputs Sheet'''
def writeSAIHead(head, final, intervallabel, wb_wt_SmallAreaInputs, col):
    for y in range(1, final + 1):
        col += 1
        wb_wt_SmallAreaInputs.cell(3, col).value = head
        wb_wt_SmallAreaInputs.cell(4, col).value = intervallabel[y]
    col += 1
    return col, wb_wt_SmallAreaInputs

'''Sub-function for writing input data to SmallAreaInputs Sheet'''
def writeSAIData(final, dataset, wb_wt_SmallAreaInputs, col, row, sex, area, age_group):
    for y in range(1, final + 1):
        col += 1
        wb_wt_SmallAreaInputs.cell(row, col).value = dataset[y, area, sex, age_group]
    col += 1
    return col, wb_wt_SmallAreaInputs

'''Function for writing input data to SmallAreaInputs Sheet'''
def writeSAI(wb_wt_SmallAreaInputs, intervallabel, numareas, lastage, final, Areacode, Areaname, pclabel, ASDR, ASOMR, inward, agelabel, ASFR):
    col = 4

    # Record Headings for SmallAreaInputs sheet
    col, wb_wt_SmallAreaInputs = writeSAIHead("F ASDRs", final, intervallabel, wb_wt_SmallAreaInputs, col)
    col, wb_wt_SmallAreaInputs = writeSAIHead("M ASDRs", final, intervallabel, wb_wt_SmallAreaInputs, col)
    col, wb_wt_SmallAreaInputs = writeSAIHead("F ASOMRs", final, intervallabel, wb_wt_SmallAreaInputs, col)
    col, wb_wt_SmallAreaInputs = writeSAIHead("M ASOMRs", final, intervallabel, wb_wt_SmallAreaInputs, col)
    col, wb_wt_SmallAreaInputs = writeSAIHead("F inward", final, intervallabel, wb_wt_SmallAreaInputs, col)
    col, wb_wt_SmallAreaInputs = writeSAIHead("M inward", final, intervallabel, wb_wt_SmallAreaInputs, col)
    col, wb_wt_SmallAreaInputs = writeSAIHead("ASFRs", final, intervallabel, wb_wt_SmallAreaInputs, col + 1)

    # Record ASDR, ASOMR, inward migration, ASFR input data
    row = 4
    for i in range(numareas):
        for pc in range(lastage + 1):

            # Record row name for each age-sex cohort input data in each area
            row += 1
            wb_wt_SmallAreaInputs.cell(row, 1).value = i
            wb_wt_SmallAreaInputs.cell(row, 2).value = Areacode[i]
            wb_wt_SmallAreaInputs.cell(row, 3).value = Areaname[i]
            wb_wt_SmallAreaInputs.cell(row, 4).value = pclabel[pc]
            col = 4

            # Record input data (ASDR, ASOMR, inward) into SmallAreaInputs sheet
            col, wb_wt_SmallAreaInputs = writeSAIData(final, ASDR, wb_wt_SmallAreaInputs, col, row, 0, i, pc)
            col, wb_wt_SmallAreaInputs = writeSAIData(final, ASDR, wb_wt_SmallAreaInputs, col, row, 1, i, pc)
            col, wb_wt_SmallAreaInputs = writeSAIData(final, ASOMR, wb_wt_SmallAreaInputs, col, row, 0, i, pc)
            col, wb_wt_SmallAreaInputs = writeSAIData(final, ASOMR, wb_wt_SmallAreaInputs, col, row, 1, i, pc)
            col, wb_wt_SmallAreaInputs = writeSAIData(final, inward, wb_wt_SmallAreaInputs, col, row, 0, i, pc)
            col, wb_wt_SmallAreaInputs = writeSAIData(final, inward, wb_wt_SmallAreaInputs, col, row, 1, i, pc)

    # Record ASFRs input data
    row = 8
    birthcol = col + 1
    for i in range(numareas):
        for a in range(7):
            # Record headings of ASFRs
            row += 1
            col = birthcol
            wb_wt_SmallAreaInputs.cell(row, birthcol).value = agelabel[a + 3]

            for y in range(1, final + 1):
                col += 1
                wb_wt_SmallAreaInputs.cell(row, col).value = ASFR[y, i, a]
        row += 11
    
    # Return the SmallAreaInputs Sheet
    return wb_wt_SmallAreaInputs

#################################################################################################################################################################
'''Function for read in National Projection Data'''
def readNatP(final, numages, sheet_NationalProjection):

    # Create empty array for storing National Projection data
    NatP = numpy.zeros((final + 1, 2, numages))

    # Load in national projection data
    row = 3
    for s in range(2):
        for a in range(numages):
            row += 1
            col = 2
            for y in range(final + 1):
                col += 1
                NatP[y, s, a] = sheet_NationalProjection.cell_value(row, col)
    
    # Return National Projection data
    return NatP

'''Function for reading in national projected birth, deaths, new migration data for projection'''
def readBDN(final, lastage, sheet_NationalProjection):
    
    # Create empty array for storing birth, death, net migration data
    NatB = numpy.zeros((final, 2))
    NatDpc = numpy.zeros((final, 2, lastage + 1))
    NatNpc = numpy.zeros((final, 2, lastage + 1))

    # Read in data
    col = 2
    for y in range(final):
        col += 1
        NatB[y, 0] = sheet_NationalProjection.cell_value(43, col)
        NatB[y, 1] = sheet_NationalProjection.cell_value(44, col)
    
    row = 47
    for s in range(2):
        for pc in range(lastage + 1):
            row += 1
            col = 2
            for y in range(final):
                col += 1
                NatDpc[y, s, pc] = sheet_NationalProjection.cell_value(row, col)
    
    row = 86
    for s in range(2):
        for pc in range(lastage + 1):
            row += 1
            col = 2
            for y in range(final):
                col += 1
                NatNpc[y, s, pc] = sheet_NationalProjection.cell_value(row, col)
    
    # Return national projected birth, deaths, net migration data
    return NatB, NatDpc, NatNpc

'''Function for writting column names in selected sheet'''
def writeCheckMD(final, lastage, intervallabel, sexlabel, pclabel, sheet):
    row = 4
    for y in range(final):
        for s in range(2):
            for pc in range(lastage + 1):
                row += 1
                sheet.cell(row, 1).value = intervallabel[y + 1]
                sheet.cell(row, 2).value = sexlabel[s]
                sheet.cell(row, 3).value = pclabel[pc]
        row += 2
    return sheet

'''Function for writing headings / row name in CheckMig, CheckDeaths, Log Sheets'''
def writeNoteCL(numareas, lastage, wb_wt_CheckMig, wb_wt_CheckDeaths, wb_wt_Log, intervallabel, sexlabel, pclabel, Areaname, final):
    
    # Column name in CheckMig Sheet
    col = 3
    for i in range(numareas):
        col += 1
        wb_wt_CheckMig.cell(2, col).value = i + 1
        wb_wt_CheckMig.cell(2, col + 3 + numareas).value = i + 1
        wb_wt_CheckMig.cell(3, col).value = Areaname[i]
        wb_wt_CheckMig.cell(3, col + 3 + numareas).value = Areaname[i]
    wb_wt_CheckMig.cell(3, col + 4 + numareas).value = "National net mig"

    # Column name in CheckDeath Sheet
    col = 3
    for i in range(numareas):
        col += 1
        wb_wt_CheckDeaths.cell(2, col).value = i + 1
        wb_wt_CheckDeaths.cell(3, col).value = Areaname[i]
    wb_wt_CheckDeaths.cell(3, col + 1).value = "National deaths"

    # Row notation / name in CheckMig Sheet
    wb_wt_CheckMig = writeCheckMD(final, lastage, intervallabel, sexlabel, pclabel, wb_wt_CheckMig)

    # Row notation / name in CheckDeath Sheet
    wb_wt_CheckDeaths = writeCheckMD(final, lastage, intervallabel, sexlabel, pclabel, wb_wt_CheckDeaths)

    # Headings in Log Sheet
    wb_wt_Log.cell(1, 1).value = "No. of iterations"
    wb_wt_Log.cell(3, 1).value = "Projection interval"
    wb_wt_Log.cell(3, 2).value = "Population-at-risk iterations"
    wb_wt_Log.cell(3, 3).value = "Mig adjustment iterations"

    return wb_wt_CheckMig, wb_wt_CheckDeaths, wb_wt_Log