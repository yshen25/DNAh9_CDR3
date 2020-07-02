#NOTE: Amino Acid = AA

from openpyxl import load_workbook

aminoacids = {'G': 0, 'A': 0, 'V': 0, 'C': 0, 'P': 0, 'L': 0,
              'I': 0, 'M': 0, 'W': 0, 'F': 0, 'K': 1, 'R': 1,
              'H': .1, 'S': 0, 'T': 0, 'Y': 0, 'N': 0, 'Q': 0,
              'D': -1, 'E': -1}

#Calculates NCPR by adding all the charges of each AA
#and dividing by the total number of AAs in the Sequence
def CalculateNCPR(Sequence):
    x = 0
    totalCharge = 0
    while x < len(Sequence):
        try:
            if x != 0:
                AA = Sequence[x-1:x]
            else:
                AA = Sequence[0]
            Charge = aminoacids[AA]
            totalCharge += Charge
        except:
            print(Sequence[x-1:x]+" is not an Amino Acid. Please check your sequence and try again")
            exit()
        x+=1
    NCPR = totalCharge/len(Sequence)
    return NCPR

#Checks ID from "metastatic and primary CDR3s b" sheet to find match in "DNAH9"
#If match is found returns Reference_Allele, Tumor_Seq_Allele1, Tumor_Seq_Allele2
def CheckID(ID,sheet2):
    for y in range(2, sheet2.max_row + 1):
        ID2 = sheet2["O{cellRow}".format(cellRow = y)].value
        if ID == ID2:
            Mutation_Type = sheet2["H{cellRow}".format(cellRow = y)].value
            if Mutation_Type != "Missense_Mutation":
                return "N/A", "N/A", "N/A"
            Reference_Allele = sheet2["J{cellRow}".format(cellRow = y)].value
            Tumor_Seq_Allele1 = sheet2["K{cellRow}".format(cellRow = y)].value
            Tumor_Seq_Allele2 = sheet2["L{cellRow}".format(cellRow = y)].value
            return Reference_Allele, Tumor_Seq_Allele1, Tumor_Seq_Allele2
    return "N/A", "N/A", "N/A"

#Calculates Change in Charge by checking the change in charge of the
#Tumor_Seq_Allele1(TAA1) and Tumor_Seq_Allele2(TAA2) against the Reference Allele(Ref_AA)
def Calc_AA_Charge_Δ(Ref_AA,TAA1,TAA2):
    if Ref_AA == "N/A":
        return "N/A"
    if TAA1 != "-":
        AA_Charge_Δ = aminoacids[TAA1]-aminoacids[Ref_AA]
    if TAA2 != "-":
        AA_Charge_Δ = aminoacids[TAA2]-aminoacids[Ref_AA]
    print(AA_Charge_Δ)
    return AA_Charge_Δ

#If AA_Charge_Δ = 0 multiplying makes NCPR_CS 0, so if  AA_Charge_Δ = 0
#AA_Charge_Δ is exluded from calculations
def Calc_NCPR_CS(AA_Charge_Δ,NCPR):
    if AA_Charge_Δ == "N/A":
        return "N/A"
    elif AA_Charge_Δ == 0:
        return NCPR * -1
    NCPR_CS = AA_Charge_Δ * NCPR * -1
    return NCPR_CS

workbook = load_workbook(filename="Dataset.xlsx")
sheet = workbook['metastatic and primary CDR3s b']
sheet2 = workbook['DNAH9']

for x in range(2, sheet.max_row + 1):
    #Grab Patient ID
    Patient_ID = sheet["A{cellRow}".format(cellRow = x)].value
    #Checks ID for Match in DHA9
    Ref_Allele, Tumor_Seq_Allele1, Tumor_Seq_Allele2 = CheckID(Patient_ID,sheet2)
    #Calculates Change in Charge of AA
    AA_Charge_Δ = Calc_AA_Charge_Δ(Ref_Allele,Tumor_Seq_Allele1,Tumor_Seq_Allele2)
    #Calculates NCPR
    NCPR = CalculateNCPR(sheet["Q{cellRow}".format(cellRow = x)].value)
    #Calculates NCPR excluding AA_Charge_Δ when AA_Charge_Δ = 0
    NCPR_CS = Calc_NCPR_CS(AA_Charge_Δ,NCPR)

    #Assigns all values to rows in new datasheet
    sheet["U{cellRow}".format(cellRow = x)] = Ref_Allele
    sheet["V{cellRow}".format(cellRow = x)] = Tumor_Seq_Allele1
    sheet["W{cellRow}".format(cellRow = x)] = Tumor_Seq_Allele2
    sheet["X{cellRow}".format(cellRow = x)] = AA_Charge_Δ
    sheet["T{cellRow}".format(cellRow = x)] = NCPR
    sheet["Y{cellRow}".format(cellRow = x)] = NCPR_CS
    
# Save the spreadsheet
workbook.save(filename="EditedDataset.xlsx")
