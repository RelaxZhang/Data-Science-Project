# Loading Packages
library("readxl")

# Read in Excel Files (Original & Output)
wt_loc = '../DSP/DS_Project/VBA_Benchmark/vba.xlsx'
rd_loc = '../DSP/DS_Project/VBA_Benchmark/original.xlsm'
wb_rd = read_excel(rd_loc)
wb_wt = read_excel(wt_loc)