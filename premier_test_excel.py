from openpyxl import*
from openpyxl.utils import*
import pandas as pd
read_file = pd.read_csv (r'C:/Users/Etudiant/Desktop/stage_excel/ExpDevis.csv',delimiter=";", encoding='latin1')
read_file.to_excel (r'C:/Users/Etudiant/Desktop/stage_excel/ExpDevis.xlsx', index = None, header=True)
export= load_workbook('ExpDevis.xlsx')
data= load_workbook('HYFR_DC-FC_2022 JZ.xlsx')
ex = export.active
da = data["Data"]
colone=["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ"]
ligne_ex=0
for row in ex:
    ligne_ex+=1
ligne_data=3
while da[colone[6]+str(ligne_data)].data_type =="s":
    ligne_data+=1
all_devis=[]
for i in range(ligne_data): 
    if type(da[colone[4]+str(i+1)].value)==int:
        all_devis.append(da[colone[4]+str(i+1)].value)

                                                            ################################ données devis ################################

Date_Devis = (ex[colone[0]+str(ligne_ex)].value)
bu_key=f"=IFERROR(INDEX(Tabelle2[BU],MATCH(tbl_DCFC[[#This Row],[Categorie]],CAT,0)),\"\")"
cw=f"=IF({'D' + str(ligne_data)}<>\"\",\"S\"&TEXT(WEEKNUM({'D' + str(ligne_data)},21),\"00\"),\"\")"
N_Devis = (ex[colone[1]+str(ligne_ex)].value)
Client=(ex[colone[2]+str(ligne_ex)].value)
ASM= (ex[colone[3]+str(ligne_ex)].value)
Montant= int(ex[colone[4]+str(ligne_ex)].value)
Date=Date_Devis
Status=None
if (ex[colone[5]+str(ligne_ex)].data_type)=="s":
    Status='Accepté'
if Montant==0:
    Status='Opportunité'
indice_devis=all_devis.count(N_Devis)
index=colone[indice_devis]
Client_end=None
Catégorie=None
données_devis=[Date,bu_key,cw,Date_Devis,N_Devis,index, Client,Client_end,Montant,ASM,Catégorie,Status]

                                                            ################################ Split ################################
split=""


                                                            ################################ données commande ################################
cw_order=f"=IF({'Q' + str(ligne_data)}<>\"\",\"S\"&TEXT(WEEKNUM({'Q' + str(ligne_data)},21),\"00\"),\"\")"
OI_Year=f"=IF(Data!${'Q'+ str(ligne_data)}=\"\",\"\",YEAR(Data!${'Q'+ str(ligne_data)}))"
OI_Month=f"=IF(Data!${'Q'+ str(ligne_data)}=\"\",\"\",MONTH(Data!${'Q'+ str(ligne_data)}))"
date_comande=(ex[colone[5]+str(ligne_ex)].value)
n_commande=(ex[colone[6]+str(ligne_ex)].value)
montant_commande=int(ex[colone[7]+str(ligne_ex)].value)
montant_restant=""
délai=""
données_commande=[split,cw_order,OI_Year,OI_Month,date_comande,n_commande,montant_commande,montant_restant,délai]
                                                            ################################ Split2 ################################
split2=""

                                                            ################################ données facture ################################
date_réelle=""
cw_rev=f"=IF({'AC'+ str(ligne_data)}<>\"\",\"S\"&TEXT(WEEKNUM({'AC'+ str(ligne_data)},21),\"00\"),\"\")"
revenue_Year=f"=IF(Data!${'AC'+ str(ligne_data)}=\"\",\"\",YEAR(Data!${'AC'+ str(ligne_data)}))"
revenue_Month=f"=IF(Data!${'AC'+ str(ligne_data)}=\"\",\"\",MONTH(Data!${'AC'+ str(ligne_data)}))"
dateliv_réelle=(ex[colone[8]+str(ligne_ex)].value)
n_facture=""
date_facture=(ex[colone[9]+str(ligne_ex)].value)
données_facture=[split2,date_réelle,cw_rev,revenue_Year,revenue_Month,dateliv_réelle,n_facture,date_facture]

                                                              ###############################  vérif devis  ################################
# verif_devis=données_devis
# verif_devis[2]=f"=IF({'D' + str(ligne_data-1)}<>\"\",\"S\"&TEXT(WEEKNUM({'D' + str(ligne_data-1)},21),\"00\"),\"\")"
# verif_devis[5]=colone[indice_devis-1]
# verif_devispre=[]
# for i in range(len(données_devis)):
#     verif_devispre.append(da[colone[i]+str(ligne_data-1)].value)
                                                            ################################ insersion excel ################################

données=données_devis+données_commande+données_facture
for i in range(len(données)):
    da[colone[i]+str(ligne_data)]=données[i]
data.save('HYFR_DC-FC_2022 JZ.xlsx')
