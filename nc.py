from openpyxl import*
from openpyxl.utils import*
import datetime as dt
colone = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA",
          "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ"]

nom=input("FNC?\n")
export = load_workbook(nom+'.xlsx', data_only=True)
monitoring = load_workbook('FR 1005-Monitoring NC & Action Plan-V001.xlsx')
checklist = load_workbook('checklist.xlsm', keep_vba=True)
mo=monitoring["Suivi_FNC"]
ex = export["Fiche de Non-Conformité"]
ch=checklist["INT_DATA"]
ch_ligne=1
mo_ligne=1
all_fnc=[]
while ch[colone[1]+str(ch_ligne)].data_type == "s":
    ch_ligne += 1
while mo[colone[0]+str(mo_ligne)].data_type == "s":
    mo_ligne += 1

for i in range(3,mo_ligne):
    all_fnc.append(mo[colone[0]+str(i)].value)
print(all_fnc)
données=[]
données.append(ex["l7"].value)
données.append (dt.datetime.strftime((ex["AJ7"].value), "%d/%m/%Y"))
données.append (str((ex["S10"].value)).upper()+' '+ (ex["F10"].value))
détection=""
if(ex["AT7"].value)==1 :
    détection="Fournisseur"
if(ex["AT7"].value)==2 :
    détection="Interne"
if(ex["AT7"].value)==3 :
    détection="Client"
données.append(détection)
données.append("")
données.append(ex["AN16"].value) 
données.append(ex["C13"].value)
données.append(ex["AH13"].value)
#####section 3 
données.append(ex["j32"].value) 
données.append(ex["p32"].value) 
données.append(ex["v32"].value) 
données.append(ex["ak32"].value)
#####section 4
solution=""
if(ex["d38"].value)!="":
    if solution=="":
        solution+="MFT "
    else:
        solution+="+ MFT "
if(ex["k40"].value)!="":
    if solution=="":
        solution+="5 Why "
    else:
        solution+="+ 5 Why "
if(ex["m50"].value)!="":
    if solution=="":
        solution+="Ishikawa "
    else:
        solution+="+ Ishikawa "
données.append(solution)
données.append(ex["d38"].value+ex["k40"].value+ex["m50"].value)
données.append(ex["au9"].value)
for i in range (len(all_fnc)):
    if nom == all_fnc[i]:
        mo_ligne=i+3

    for i in range(len(données)):
        mo[colone[i]+str(mo_ligne)] = données[i]

ch[colone[1]+str(ch_ligne)] = données[0]
monitoring.save('FR 1005-Monitoring NC & Action Plan-V001.xlsx')
checklist.save('test.xlsm')