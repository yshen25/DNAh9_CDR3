#!usr/bin/env python3

from openpyxl import load_workbook
from SupportFunctions import *

workbook = load_workbook(filename="Dataset.xlsx")
CDR3File = workbook['metastatic and primary CDR3s b']
DNAH9 = workbook['DNAH9']
Clinical = workbook['clinical']
PatientSurvivalChart = workbook['PatientSurvivalChart']

Major_Dictionary = {}

'''
Dictionary Example:
    {
        "TCGA-D3-A1Q4":{ "TRA CDR3":[abc,abc], "TRB CDR3":[abc,abc],
        "Mutations": ["S1234D"], "Months Left": 10, "Complimentary": True
        }
    }

'''

for x in range(2, DNAH9.max_row + 1):
    Patient_ID = DNAH9["A{cellRow}".format(cellRow = x)].value
    if Patient_ID in Major_Dictionary:
        Mutation = DNAH9["E{cellRow}".format(cellRow = x)].value
        Major_Dictionary[Patient_ID]["Mutations"].append(Mutation)
    else:
        Major_Dictionary[Patient_ID] = {"TRA CDR3":[], "TRB CDR3":[],
        "Mutations": [], "Months Left": False, "Complimentary": False}
        Mutation = DNAH9["E{cellRow}".format(cellRow = x)].value
        Major_Dictionary[Patient_ID]["Mutations"].append(Mutation)
        
for x in range(2, CDR3File.max_row + 1):
    Patient_ID = CDR3File["A{cellRow}".format(cellRow = x)].value
    if Patient_ID in Major_Dictionary:
        CDR3 = CDR3File["B{cellRow}".format(cellRow = x)].value
        Receptor = CDR3File["C{cellRow}".format(cellRow = x)].value

        Major_Dictionary[Patient_ID]["{receptor} CDR3".format(receptor = Receptor)].append(CDR3)

        MonthsLeft = MonthsLeftFunc(Patient_ID)
        Major_Dictionary[Patient_ID]["Months Left"] = MonthsLeft
        
for Patient_ID in Major_Dictionary:
    ID_Dictionary = Major_Dictionary[Patient_ID]
    
    for Mutation in Major_Dictionary[Patient_ID]["Mutations"]:
        MutationSide = LeftorRight(Mutation[1:len(Mutation)-1])
        if MutationSide == "Left":
            for CDR3 in ID_Dictionary["TRA CDR3"]:
                if ComplimentaryFunction(CDR3,Mutation) == True:
                    Major_Dictionary[Patient_ID]["Complimentary"] = True
                    break
        if MutationSide == "Right":
            for CDR3 in ID_Dictionary["TRB CDR3"]:
                if ComplimentaryFunction(CDR3,Mutation) == True:
                    Major_Dictionary[Patient_ID]["Complimentary"] = True
                    break
        if MutationSide == "Both":
            for CDR3 in ID_Dictionary["TRA CDR3"]:
                if ComplimentaryFunction(CDR3,Mutation) == True:
                    Major_Dictionary[Patient_ID]["Complimentary"] = True
                    break
            for CDR3 in ID_Dictionary["TRB CDR3"]:
                if ComplimentaryFunction(CDR3,Mutation) == True:
                    Major_Dictionary[Patient_ID]["Complimentary"] = True
                    break
        if MutationSide == "N/A":
            print("No Peptides", Patient_ID, Mutation)
print(Major_Dictionary)





