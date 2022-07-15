### Function Package for Generating Data for Synthetic Migration Model from 1996 - 2016 ###
import numpy as np
import pandas as pd
import math

'''Function for selecting target year for the Synthetic Migration Model from age-sex cohort data'''
def selectSynYear(all_area_pop_full, year):
    all_area_pop_select = all_area_pop_full.drop(all_area_pop_full[all_area_pop_full.Year < year].index)
    all_area_pop_select = all_area_pop_select.drop(all_area_pop_select[all_area_pop_select.Year > year].index)
    all_area_pop_select = all_area_pop_select.reset_index()
    all_area_pop_select.drop('index', inplace = True, axis = 1)
    return all_area_pop_select

'''Function for aggregating remainder areas'''
def remainderFunc(all_area_pop, remainder_list, outsider_list):
    remainder = 0
    for i in range(len(all_area_pop["Year"])):
        if (all_area_pop.iloc[i, 38] not in outsider_list):
            if (all_area_pop.iloc[i, 38] in remainder_list):
                remainder += all_area_pop.iloc[i, 36]
    return remainder

'''Function for collecting total population data from age-sex cohort dataframe'''
def totPop(remainderData, all_area_pop_year, remainder_list):
    SA3_Name_list = []
    SA3_Code_list = []
    SA3_totPop_list = []
    for i in range(len(all_area_pop_year["Year"])):
        if (all_area_pop_year.iloc[i, 38] not in remainder_list):
            SA3_Name_list.append(all_area_pop_year.iloc[i, 38])
            SA3_Code_list.append(all_area_pop_year.iloc[i, 39])
            SA3_totPop_list.append(all_area_pop_year.iloc[i, 36])
    SA3_Name_list.append("Remainder")
    SA3_Code_list.append(99999)
    SA3_totPop_list.append(remainderData)
    return SA3_Name_list, SA3_Code_list, SA3_totPop_list

'''Function for converting missing value .. in remainder areas to 0'''
def convertMissing(remain_list):
    for i in range(len(remain_list)):
        for j in range(len(remain_list[i])):
            if remain_list[i][j] == "..":
                remain_list[i][j] = 0
    return remain_list

'''Function for obtaining the age-sex ERP data for the Synthetic Migration Model'''
def agesexPop(remainder_list, all_area_pop_year, outsider_listm, numages):
    male_list = []
    female_list = []
    remain_male_list = []
    remain_female_list = []

    for i in range(len(all_area_pop_year["Year"])):
        if all_area_pop_year.iloc[i, 38] not in remainder_list:
            male_list.append(all_area_pop_year.iloc[i].tolist()[0:numages])
            female_list.append(all_area_pop_year.iloc[i].tolist()[numages: (2 * numages)])
        else:
            if all_area_pop_year.iloc[i, 38] not in outsider_listm:
                remain_male_list.append(all_area_pop_year.iloc[i].tolist()[0:numages])
                remain_female_list.append(all_area_pop_year.iloc[i].tolist()[numages: (2 * numages)])

    remain_male_list = convertMissing(remain_male_list)
    remain_female_list = convertMissing(remain_female_list)

    male_remainder = np.zeros(numages)
    female_remainder = np.zeros(numages)

    for i in range(len(remain_male_list)):
        male_remainder += np.array(remain_male_list[i])
        female_remainder += np.array(remain_female_list[i])

    male_list.append(male_remainder.tolist())
    female_list.append(female_remainder.tolist())

    return male_list, female_list

'''Function for recording Age-Sex ERP data into Worksheet'''
def recordAgeSex(male_list_before, male_list_jump, female_list_before, female_list_jump, wb_wt_AgeSexERPs):
    row = 4
    for i in range(len(male_list_before)):
        for j in range(len(male_list_before[i])):
            row += 1
            wb_wt_AgeSexERPs.cell(row, 5).value = male_list_before[i][j]
            wb_wt_AgeSexERPs.cell(row, 6).value = male_list_jump[i][j]
            wb_wt_AgeSexERPs.cell(row, 7).value = female_list_before[i][j]
            wb_wt_AgeSexERPs.cell(row, 8).value = female_list_jump[i][j]
    return wb_wt_AgeSexERPs

####################################################################################################################

'''Function for calculating annual average growth & growth rate of population in each area'''
def growthRate(previous, jump, df_ERPs):
    growth = []
    growth_rate = []
    for i in range(len(df_ERPs[str(previous)])):
        if (df_ERPs[str(previous)][i] <= 0):
            growth.append(growth)
            growth_rate.append(0)
        else:
            growth.append((df_ERPs[str(jump)][i] - df_ERPs[str(previous)][i]) / (jump - previous))
            growth_rate.append(math.log(df_ERPs[str(jump)][i] / df_ERPs[str(previous)][i]) / (jump - previous))
    return growth, growth_rate

'''Generalised CSP_model Function'''
def CSP(ERP_list, target_NPT, column_name):
    
    # Generate Share of Population for CSP_model
    share_of_pop = []
    for i in ERP_list:
        share_of_pop.append(i / sum(ERP_list))

    # Generate total population data in each region based on the CSP_model's logic
    CSP_model = []
    for i in share_of_pop:
        each_area = []
        for j in target_NPT:
            each_area.append(i * j)
        CSP_model.append(each_area) 
    
    # Create the dataframe for storing CSP information and return the value
    df_CSP = pd.DataFrame(CSP_model, columns = column_name)
    return df_CSP

'''Function for calculating Population Ceiling K list for MEX_Model'''
def popCeil(jumpoff, coef, growth_rate, df_ERPs):
    pop_ceil = []
    for i in range(len(df_ERPs[str(jumpoff)])):
        if growth_rate[i] >= 0:
            pop_ceil.append(coef * df_ERPs[str(jumpoff)][i])
        else:
            pop_ceil.append(1 / coef * df_ERPs[str(jumpoff)][i])
    return pop_ceil

'''Function for generating total population with the MEX_Model'''
def MEX(jumpoff, target_final, df_ERPs, pop_ceil, growth_rate):

    # Prepare dataframe for storing MEX_Model result
    gap_process = target_final - jumpoff
    base_year = df_ERPs[str(jumpoff)].tolist()
    df_MEX = {str(jumpoff): base_year}
    df_MEX = pd.DataFrame(df_MEX)

    # Record projection data year by year
    for i in range(gap_process):
        project_year = []
        for j in range(len(base_year)):
            if (growth_rate[j] >= 0):
                project_year.append(base_year[j] * math.exp(growth_rate[j] * (1 - (base_year[j] / pop_ceil[j]))))
            else:
                project_year.append(base_year[j] * math.exp(growth_rate[j] * (1 - (pop_ceil[j] / base_year[j]))))
        df_MEX[str(jumpoff + 1 + i)] = project_year
        base_year = project_year
    
    # Return MEX_Model Result
    return df_MEX