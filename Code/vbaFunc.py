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
def readERP(intervals, numareas, numages, sheet_agesex):

    # Create empty array for storing ERP in small area by age and sex
    infoERP = numpy.zeros((intervals, numareas, 2, numages))
    infoERP[:] = numpy.nan

    # Collect data
    for year in range(intervals):
        row = 3
        for i in range(numareas):
            for a in range(numages):
                row += 1
                # Females
                infoERP[year, i, 0, a] = sheet_agesex.cell_value(row, 6 + year)
                # Males
                infoERP[year, i, 1, a] = sheet_agesex.cell_value(row, 4 + year)
    
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
            tempbirths[a] = ASFR_data[0, i, a] * 2.5 * (ERP[0, i, 0, a] + ERP[1, i, 0, a])
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
