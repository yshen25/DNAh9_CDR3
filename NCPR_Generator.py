#!usr/bin/env python3
#NOTE: Amino Acid = AA

from openpyxl import load_workbook
from SupportFunctions import *

workbook = load_workbook(filename="Dataset.xlsx")
sheet = workbook['metastatic and primary CDR3s b']
sheet2 = workbook['DNAH9']
sheet3 = workbook['clinical']
sheet4 = workbook['PatientSurvivalChart']

for x in range(2, sheet.max_row + 1):
    #Grab Patient ID
    Patient_ID = sheet["A{cellRow}".format(cellRow = x)].value
    #Checks ID for Match in DHA9
    Ref_Allele, Tumor_Seq_Allele = CheckID(Patient_ID,sheet2)
    #Calculates Change in Charge of AA
    AA_Charge_Δ = Calc_AA_Charge_Δ(Ref_Allele,Tumor_Seq_Allele)
    #Finds days to death
    MonthsLeft = MonthsLeftFunc(Patient_ID,sheet3)
    #Calculates NCPR
    NCPR = CalculateNCPR(sheet["B{cellRow}".format(cellRow = x)].value)
    #Calculates NCPR excluding AA_Charge_Δ when AA_Charge_Δ = 0
    NCPR_CS = Calc_NCPR_CS(AA_Charge_Δ,NCPR,Patient_ID)
    #Calculate Complimentary Value
    if sheet["J{cellRow}".format(cellRow = x)].value != True:
        if NCPR_CS != "N/A" and NCPR_CS > 0:
            MakeTrue(Patient_ID,sheet)
        elif NCPR_CS != "N/A":
            sheet["J{cellRow}".format(cellRow = x)] = False
        else:
            sheet["J{cellRow}".format(cellRow = x)] = "N/A"
    #Assigns all values to rows in new datasheet
    sheet["E{cellRow}".format(cellRow = x)] = Ref_Allele
    sheet["F{cellRow}".format(cellRow = x)] = Tumor_Seq_Allele
    sheet["G{cellRow}".format(cellRow = x)] = AA_Charge_Δ
    sheet["D{cellRow}".format(cellRow = x)] = NCPR
    sheet["H{cellRow}".format(cellRow = x)] = NCPR_CS
    sheet["I{cellRow}".format(cellRow = x)] = MonthsLeft
    
# Save the spreadsheet
workbook.save(filename="EditedDataset.xlsx")

# Reopens SpreadSheet to use saved Complimentary Values
workbook = load_workbook(filename="EditedDataset.xlsx")
sheet = workbook['metastatic and primary CDR3s b']
sheet4 = workbook['PatientSurvivalChart']

#Makes survival Chart
PercentAlive(sheet,sheet4)

#81 NonCom Patients
#64 Comp Patients

#Saves Workbook
workbook.save(filename="EditedDataset2.xlsx")
