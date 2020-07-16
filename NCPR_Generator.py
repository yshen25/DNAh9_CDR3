#!usr/bin/env python3
#NOTE: Amino Acid = AA

from openpyxl import load_workbook

"""AliveStillComp = {}
AliveStillNon = {}

x = 0
while x < 501:
    AliveStillComp[x] = 0
    AliveStillNon[x] = 0
    x+=1"""

def get_KD_original():
    """ Function which returns the original KD hydropathy lookup table
    """    
    
    return  {'I': 4.5,                
             'V': 4.2,
             'L': 3.8,
             'F': 2.8,
             'C': 2.5,
             'M': 1.9,
             'A': 1.8,
             'G': -0.4,
             'T': -0.7,
             'S': -0.8,
             'W': -0.9,
             'Y': -1.3,
             'P': -1.6,
             'H': -3.2,
             'E': -3.5,
             'Q': -3.5,
             'D': -3.5,
             'N': -3.5,
             'K': -3.9,
             'R': -4.5}

def get_AA_charge(AA):
    """ Function which returns the original KD hydropathy lookup table
    """    
    
    AA_Charge_Table = {'I': 0,                
             'V': 0,
             'L': 0,
             'F': 0,
             'C': 0,
             'M': 0,
             'A': 0,
             'G': 0,
             'T': 0,
             'S': 0,
             'W': 0,
             'Y': 0,
             'P': 0,
             'H': .1,
             'E': -1,
             'Q': 0,
             'D': -1,
             'N': 0,
             'K': 1,
             'R': 1}
    return AA_Charge_Table[AA]

#Calculates NCPR by adding all the charges of each AA
#and dividing by the total number of AAs in the Sequence
def CalculateNCPR(Sequence):
    totalCharge = 0
    for AA in Sequence:
        Charge = get_AA_charge(AA)
        totalCharge += Charge
    NCPR = totalCharge/len(Sequence)
    return NCPR
#Checks ID from "metastatic and primary CDR3s b" sheet to find match in "DNAH9"
#If match is found returns Reference_Allele, Tumor_Seq_Allele1, Tumor_Seq_Allele2
def CheckID(ID,sheet3):
    for y in range(2, sheet2.max_row + 1):
        ID2 = sheet2["A{cellRow}".format(cellRow = y)].value
        if ID == ID2:
            AminoAcids = sheet2["C{cellRow}".format(cellRow = y)].value
            Reference_Allele = AminoAcids[0]
            Tumor_Seq_Allele = AminoAcids[2]
            return Reference_Allele, Tumor_Seq_Allele
    return "N/A", "N/A"

def aliveStillComp(MonthsLeft):
    AliveStillComp[MonthsLeft] = AliveStillComp[MonthsLeft] - 1
def aliveStillNon(MonthsLeft):
    AliveStillNon[MonthsLeft] = AliveStillNon[MonthsLeft] - 1

def MonthsLeftFunc(ID,sheet3):
    for y in range(2, sheet3.max_row + 1):
        ID2 = sheet3["B{cellRow}".format(cellRow = y)].value
        if ID == ID2:
            daysLeft = sheet3["J{cellRow}".format(cellRow = y)].value
            try:
                return int(daysLeft/30)
            except:
                return daysLeft
    return "N/A"

#Calculates Change in Charge by checking the change in charge of the
#Tumor_Seq_Allele1(TAA1) and Tumor_Seq_Allele2(TAA2) against the Reference Allele(Ref_AA)
def Calc_AA_Charge_Δ(Ref_AA,TAA):
    if Ref_AA == "N/A":
        return "N/A"
    if TAA != "-":
        AA_Charge_Δ = get_AA_charge(TAA)-get_AA_charge(Ref_AA)
    return AA_Charge_Δ

#If AA_Charge_Δ = 0 multiplying makes NCPR_CS 0, so if  AA_Charge_Δ = 0
#AA_Charge_Δ is exluded from calculations

def Calc_NCPR_CS(AA_Charge_Δ,NCPR):
    '''A positive value denotes a complementary score'''
    if AA_Charge_Δ == "N/A":
        return "N/A"
    elif AA_Charge_Δ == 0:
        return NCPR * -1
    NCPR_CS = AA_Charge_Δ * NCPR * -1
    return NCPR_CS

def MakeTrue(Patient_ID,sheet):
    for x in range(2, sheet.max_row + 1):
        if Patient_ID == sheet["A{cellRow}".format(cellRow = x)].value:
            sheet["J{cellRow}".format(cellRow = x)] = True

"""def PercentAlive(sheet,sheet4):
    x = 2
    while x < sheet.max_row:
        Patient_ID = sheet["A{cellRow}".format(cellRow = x)].value
        MonthsLeft = sheet["I{cellRow}".format(cellRow = x)].value
        Complimentary = sheet["J{cellRow}".format(cellRow = x)].value
        while Patient_ID == sheet["A{cellRow}".format(cellRow = x)].value:
            x+=1
        if MonthsLeft != "'--":
            if Complimentary == True:
                AliveStillComp[MonthsLeft] = AliveStillComp[MonthsLeft] + 1
            else:
                AliveStillNon[MonthsLeft] = AliveStillNon[MonthsLeft] + 1
    x = 3
    CompCurrent = 315
    NonCompCurrent = 315

    while x <= 500:
        CompCurrent -= AliveStillComp[x]
        NonCompCurrent -= AliveStillNon[x]
        sheet4["B{cellRow}".format(cellRow = x)] = CompCurrent/315*100
        sheet4["D{cellRow}".format(cellRow = x)] = NonCompCurrent/315*100
        x += 1"""

workbook = load_workbook(filename="Dataset.xlsx")
sheet = workbook['metastatic and primary CDR3s b']
sheet2 = workbook['DNAH9']
sheet3 = workbook['clinical']

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
    NCPR_CS = Calc_NCPR_CS(AA_Charge_Δ,NCPR)
    #Calculate Complimentary Value
    if sheet["J{cellRow}".format(cellRow = x)].value != True:
        if NCPR_CS != "N/A" and NCPR_CS > 0:
            MakeTrue(Patient_ID,sheet)
        else:
            sheet["J{cellRow}".format(cellRow = x)] = False
    #Assigns all values to rows in new datasheet
    sheet["E{cellRow}".format(cellRow = x)] = Ref_Allele
    sheet["F{cellRow}".format(cellRow = x)] = Tumor_Seq_Allele
    sheet["G{cellRow}".format(cellRow = x)] = AA_Charge_Δ
    sheet["D{cellRow}".format(cellRow = x)] = NCPR
    sheet["H{cellRow}".format(cellRow = x)] = NCPR_CS
    sheet["I{cellRow}".format(cellRow = x)] = MonthsLeft
    
# Save the spreadsheet
workbook.save(filename="EditedDataset.xlsx")

"""workbook = load_workbook(filename="EditedDataset.xlsx")
sheet = workbook['metastatic and primary CDR3s b']
sheet4 = workbook['PatientSurvivalChart']

PercentAlive(sheet,sheet4)

workbook.save(filename="EditedDataset2.xlsx")"""
