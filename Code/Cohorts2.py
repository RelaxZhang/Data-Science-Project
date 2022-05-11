import pandas as pd
table = pd.read_excel("Historic time series_age5_sex_1981-2011.xls",sheet_name="Table 1",header = 4)

#NAs here are just formatting and not actual data
table = table.dropna()

#table.to_csv("Sample.csv")

allint = table.loc[~((table.iloc[:,10:] == "..").any(axis=1))]
allint = allint.reset_index()
nulls = table.loc[((table.iloc[:,10:] == "..").any(axis=1))]
nulls = nulls.reset_index()

allintgroups = table.loc[~((table.iloc[:,10:] == "..").any(axis=1))].loc[:,'SA3 Name']
nullsgroups = table.loc[((table.iloc[:,10:] == "..").any(axis=1))].loc[:,'SA3 Name']
below1000groups = allint[allint['Total'] < 1000].loc[:,'SA3 Name']
nullsgroups.to_csv("nullsgroups.csv")
below1000groups.to_csv("below1000groups.csv")

above1000 = table.loc[~(table.loc[:,'SA3 Name'].isin(below1000groups) | table.loc[:,'SA3 Name'].isin(nullsgroups)),:]
below1000 = table.loc[table.loc[:,'SA3 Name'].isin(below1000groups) | table.loc[:,'SA3 Name'].isin(nullsgroups),:]

above1000.to_csv("above1000.csv")
below1000.to_csv("below1000.csv")
