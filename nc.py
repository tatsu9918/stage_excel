from openpyxl import*
from openpyxl.utils import*
colone = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA",
          "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ"]

nom=input("FNC?\n")
export = load_workbook(nom+'.xlsx', data_only=True)
monitoring = load_workbook(nom+'.xlsx')
mo=monitoring["Suivi_FNC"]
ex = export["Fiche de Non-Conformité"]
ligne_mo=1
while mo[colone[0]+str(ligne_mo)].data_type == "s":
    ligne_mo += 1
print(ex["l7"].value)
print(ex["AJ7"].value)
print(ex["F10"].value)
print(ex["S10"].value)
détection=""
if(ex["AT7"].value)==1 :
    détection="Fournisseur"
if(ex["AT7"].value)==2 :
    détection="Interne"
if(ex["AT7"].value)==3 :
    détection="Client"
print(détection)
print("")
print(ex["AN16"].value) 
print(ex["C13"].value)
print(ex["AH13"].value) 
#####section 3 
print(ex["j32"].value) 
print(ex["p32"].value) 
print(ex["v32"].value) 
print(ex["ak32"].value)
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
print(solution)
print(ex["d38"].value,ex["k40"].value,ex["m50"].value)
print(ex["au9"].value)