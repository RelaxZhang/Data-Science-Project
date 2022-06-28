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
        male04 = jumpoffERP[i][1][0]
        female04 = jumpoffERP[i][0][0]
        female_1519 = jumpoffERP[i][0][3]
        female_2024 = jumpoffERP[i][0][4]
        female_2529 = jumpoffERP[i][0][5]
        female_3034 = jumpoffERP[i][0][6]
        female_3539 = jumpoffERP[i][0][7]
        female_4044 = jumpoffERP[i][0][8]
        female_4549 = jumpoffERP[i][0][9]
        
        # Subjective definition for simplifying formula
        new_born = male04 + female04
        most_fert = female_2529 + female_3034
        all_fert = female_1519 + female_2024 + female_2529 + female_3034 + female_3539 + female_4044 + female_4549

        # Calculate result for xTFR in each area
        result_xTFR[i] = (10.65 - (12.55 * most_fert / all_fert)) * new_born / all_fert

    # Return xTRF result for storing into Fertility Sheet
    return result_xTFR

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

    # Create list for storing Estimated & Projected Small Area Total Population Results
    smallAreaTotalPop = [[None] * intervals] * numareas
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
            smallAreaTotalPop[a][i] = sheet_SmallAreaTotals.cell_value(row, SATP_col)
    
    # Return Small Area Total Population for specific range of years
    return smallAreaTotalPop