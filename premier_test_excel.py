from openpyxl import Workbook, load_workbook
import pandas as pd
read_file = pd.read_csv (r'C:/Users/Etudiant/Desktop/stage_excel/ExpDevis.csv',delimiter=";", encoding='latin1')
read_file.to_excel (r'C:/Users/Etudiant/Desktop/stage_excel/ExpDevis.xlsx', index = None, header=True)
export= load_workbook('ExpDevis.xlsx')
data= load_workbook('HYFR_DC-FC_2022 JZ.xlsx')
ex = export.active
da = data["Data"]
colone=["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ"]
# ligne_ex=0
# for row in ex:
#     ligne_ex+=1
# Date_Devis = (ex[colone[0]+str(ligne_ex)].value)
# N_Devis = (ex[colone[1]+str(ligne_ex)].value)
# Client=(ex[colone[2]+str(ligne_ex)].value)
# ASM= (ex[colone[3]+str(ligne_ex)].value)
# Montant= (ex[colone[4]+str(ligne_ex)].value)
# Date=Date_Devis
da.delete_rows(1)
ligne_data=0
for row in da:
    ligne_data+=1
print(ligne_data)
# print(Date_Devis,Date,N_Devis, Client,Montant,ASM)