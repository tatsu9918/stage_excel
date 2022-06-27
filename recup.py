import os
import sys
import selenium
from selenium import webdriver
# def resource_path(relative_path):
#     try:
#         base_path = sys._MEIPASS
#     except Exception:
#         base_path = os.path.dirname(__file__)
#     return os.path.join(base_path, relative_path)
# driver = webdriver.Chrome(resource_path("chromedriver.exe"))
# list_mots = [item for item in input("Enter the list items : ").split()]
list_mots = "avion","dockage","levage","plateforme","Rafale", "C130","HERCULES","Mirage","GSE", "outillage","hélicoptere","verin","tripode","cric","portique","essais","périodique","Aibus","A330","docks","A400M", "verin tripode", "calibraiton","moyen de levage","moteur", "AIA","Clermont Ferrand"
options = webdriver.ChromeOptions()
driver = webdriver.Chrome('chromedriver.exe')
driver.get("https://www.achats.defense.gouv.fr/")
for j in range(len(list_mots)):
    barre=driver.find_element_by_id("q")
    barre.clear()
    barre.send_keys(list_mots[j])
    barre.send_keys("\n")
    offres=[]
    list_results=driver.find_elements_by_xpath("//*[@id=\"list-results\"]/div/div[1]/div/div/div[2]/h5")
    list_details=driver.find_elements_by_xpath("//*[@id=\"list-results\"]/div/div[2]/ul/li[1]/div/div[2]/span")
    list_liens=driver.find_elements_by_xpath("//*[@id=\"list-results\"]/div/div[5]/div/a[2]")
    for i in range (len(list_results)):
        if list_mots[j] in list_results[i].text or list_mots[j] in list_details[i].text:
            print(list_mots[j])
            print("")
            print(list_results[i].text)
            print(list_details[i].text)
            print(list_liens[i].get_attribute('href'))
            print("")
            
