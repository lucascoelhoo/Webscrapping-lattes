from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import os
import codecs
import unicodedata
import pandas as pd
import csv
import xlrd
import traceback

def strip_accents(text):

    try:
        text = unicode(text, 'utf-8')
    except NameError: # unicode is a default on python 3 
        pass

    text = unicodedata.normalize('NFD', text)\
           .encode('ascii', 'ignore')\
           .decode("utf-8")

    return str(text)



filename="4-nomes-coordenadores-apos-conferencia-manual\\Professores-Laboratorios-Manual.xlsx"
filename2="5-log-extracao-curriculos\\log.csv"

path_curriculos="5-curriculos-extraidos-html/"

lista_oficial = pd.read_excel(filename)
coordenadores=lista_oficial['Nome Oficial']
#laboratorios=lista_geral['NOME']
confiaveis=lista_oficial['BATEU']
coordenadores_secundarios=lista_oficial['Nome do coordenador']

try:
    os.remove(filename2)
except:{}


# tirar duplicadas de lista: list( dict.fromkeys(mylist) )
cont=0
for coordenador,confiavel,coordenador_secundario in zip(coordenadores,confiaveis,coordenadores_secundarios):
    coord_aux=""
  #  try:
    parentElement=0
    element=0
    if not(str(coordenador)=="nan"):          
        if(str(strip_accents(str(confiavel)))=="nan"):
            parentElement=0
            element=0
            try:
                # To prevent download dialog
                profile = webdriver.FirefoxProfile()
                profile.set_preference('browser.download.folderList', 2) # custom location
                profile.set_preference('browser.download.manager.showWhenStarting', False)
                profile.set_preference('browser.download.dir', '/tmp')
                profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'text/csv')
                browser = webdriver.Firefox(profile)
                #browser = webdriver.Firefox()
                browser.get('http://buscatextual.cnpq.br/buscatextual/busca.do?metodo=apresentar')
                elem = browser.find_element_by_name('textoBusca')  # Find the search box
                elem.send_keys(strip_accents(coordenador) + Keys.RETURN)
                time.sleep(15)
                #browser.findElement(By.Xpath(".//ul[@id='resultado']//[text()='Claudia Jacy Barenco Abbas']")).click();
                #browser.find_element_by_xpath('Claudia Jacy Barenco Abbas').click()
                try:
                    parentElement = browser.find_element_by_class_name("tit_form")
                    element= str(parentElement.find_elements_by_tag_name("b")[0].text)
                    #print(elementList[0].text)
                    browser.find_element_by_class_name("resultado").find_element_by_tag_name("ol").find_element_by_tag_name("li").find_element_by_tag_name("b").find_element_by_tag_name("a").click()
                    coord_aux=str(browser.find_element_by_class_name("resultado").find_element_by_tag_name("ol").find_element_by_tag_name("li").find_element_by_tag_name("b").find_element_by_tag_name("a").text)
                    #browser.find_element_by_partial_link_text(coordenador).click()
                    time.sleep(15)
                    browser.find_element_by_link_text("Abrir Currículo").click()
                    time.sleep(15)
                except:
                    parentElement = browser.find_element_by_class_name("tit_form")
                    element= str(parentElement.find_elements_by_tag_name("b")[0].text)
                    #print(elementList[0].text)
                    browser.find_element_by_class_name("resultado").find_element_by_tag_name("ol").find_element_by_tag_name("li").find_element_by_tag_name("b").find_element_by_tag_name("a").click()
                    coord_aux=str(browser.find_element_by_class_name("resultado").find_element_by_tag_name("ol").find_element_by_tag_name("li").find_element_by_tag_name("b").find_element_by_tag_name("a").text)
                    #browser.find_element_by_class_name("resultado").find_element_by_tag_name("ol").find_element_by_tag_name("li")[0].click()
                    #browser.find_element_by_partial_link_text(strip_accents(coordenador)).click()
                    time.sleep(15)
                    browser.find_element_by_link_text("Abrir Currículo").click()
                    time.sleep(15) 
                currentWindow = browser.current_window_handle 
                #store current window for backup to switch back
                for handle in browser.window_handles:
                    if handle != currentWindow:
                        browser.switch_to.window(handle)
                #now you can do your stuff in new window
                print(browser.title)
                #print(browser.find_element_by_xpath("//div[@class='menu-header']/span[@class='linksFlow']/a"))
                #browser.find_element_by_xpath("//div[@class='menu-header']/span[@class='linksFlow']/a[@class='bt-menu-header titBottom fontMenos icons-top icons-top-xml']").click()
                #time.sleep(15)
                completeName = os.path.join(path_curriculos, coordenador+".html")
                file_object = codecs.open(completeName, "w","utf-8")
                source = strip_accents(str(browser.page_source))
                #print(browser.page_source)
                html = source
                file_object.write(html)
                file_object.close()
                #with open("curriculos/"+pesquisador+".html", "w", encoding="utf-8") as f:
                #    f.write(str(browser.page_source))
                #    f.close()
                #now close the new window after doing all stuff
                browser.close()
                #after doing all stuff it new window need to switch back on main window
                browser.switch_to.window(currentWindow)
                browser.close()
                browser.quit()
                with open(filename2, 'a', newline='') as f:
                    writer = csv.writer(f)
                    if(int(element)<=1):
                        writer.writerow([coordenador,coord_aux,"SUCESSO"])
                    else:
                        writer.writerow([coordenador,coord_aux,"SUCESSO",element])
                    f.close()
                

            except:
                print("EXCECAO"+str(traceback.format_exc()))
                
                if browser and profile:
                    browser.quit()
                with open(filename2, 'a', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow([coordenador,coord_aux,"fracasso"])
                    f.close()
        else:
            parentElement=0
            element=0
            try:
                # To prevent download dialog
                profile = webdriver.FirefoxProfile()
                profile.set_preference('browser.download.folderList', 2) # custom location
                profile.set_preference('browser.download.manager.showWhenStarting', False)
                profile.set_preference('browser.download.dir', '/tmp')
                profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'text/csv')
                browser = webdriver.Firefox(profile)
                #browser = webdriver.Firefox()
                browser.get('http://buscatextual.cnpq.br/buscatextual/busca.do?metodo=apresentar')
                elem = browser.find_element_by_name('textoBusca')  # Find the search box
                elem.send_keys(strip_accents(coordenador_secundario) + Keys.RETURN)
                time.sleep(15)
                #browser.findElement(By.Xpath(".//ul[@id='resultado']//[text()='Claudia Jacy Barenco Abbas']")).click();
                #browser.find_element_by_xpath('Claudia Jacy Barenco Abbas').click()
                try:
                    parentElement = browser.find_element_by_class_name("tit_form")
                    element= str(parentElement.find_elements_by_tag_name("b")[0].text)
                    #print(elementList[0].text)
                    browser.find_element_by_class_name("resultado").find_element_by_tag_name("ol").find_element_by_tag_name("li").find_element_by_tag_name("b").find_element_by_tag_name("a").click()
                    coord_aux=str(browser.find_element_by_class_name("resultado").find_element_by_tag_name("ol").find_element_by_tag_name("li").find_element_by_tag_name("b").find_element_by_tag_name("a").text)
                    #browser.find_element_by_partial_link_text(coordenador_secundario).click()
                    time.sleep(15)
                    browser.find_element_by_link_text("Abrir Currículo").click()
                    time.sleep(15)
                except:
                    parentElement = browser.find_element_by_class_name("tit_form")
                    element= str(parentElement.find_elements_by_tag_name("b")[0].text)
                    #print(elementList[0].text)
                    browser.find_element_by_class_name("resultado").find_element_by_tag_name("ol").find_element_by_tag_name("li").find_element_by_tag_name("b").find_element_by_tag_name("a").click()
                    coord_aux=str(browser.find_element_by_class_name("resultado").find_element_by_tag_name("ol").find_element_by_tag_name("li").find_element_by_tag_name("b").find_element_by_tag_name("a").text)
                    #browser.find_element_by_partial_link_text(strip_accents(coordenador_secundario)).click()
                    time.sleep(15)
                    browser.find_element_by_link_text("Abrir Currículo").click()
                    time.sleep(15) 
                currentWindow = browser.current_window_handle 
                #store current window for backup to switch back
                for handle in browser.window_handles:
                    if handle != currentWindow:
                        browser.switch_to.window(handle)
                #now you can do your stuff in new window
                print(browser.title)
                #print(browser.find_element_by_xpath("//div[@class='menu-header']/span[@class='linksFlow']/a"))
                #browser.find_element_by_xpath("//div[@class='menu-header']/span[@class='linksFlow']/a[@class='bt-menu-header titBottom fontMenos icons-top icons-top-xml']").click()
                #time.sleep(15)
                completeName = os.path.join(path_curriculos, coordenador_secundario+".html")
                file_object = codecs.open(completeName, "w","utf-8")
                source = strip_accents(str(browser.page_source))
                #print(browser.page_source)
                html = source
                file_object.write(html)
                file_object.close()
                #with open("curriculos/"+pesquisador+".html", "w", encoding="utf-8") as f:
                #    f.write(str(browser.page_source))
                #    f.close()
                #now close the new window after doing all stuff
                browser.close()
                #after doing all stuff it new window need to switch back on main window
                browser.switch_to.window(currentWindow)
                browser.close()
                browser.quit()
                with open(filename2, 'a', newline='') as f:
                    writer = csv.writer(f)
                    if(int(element)<=1):
                        writer.writerow([coordenador_secundario,coord_aux,"SUCESSO-SECUNDARIO"])
                    else:
                        writer.writerow([coordenador_secundario,coord_aux,"SUCESSO-SECUNDARIO",element])
                    f.close()
                

            except:
                print("EXCECAO"+str(traceback.format_exc()))
                
                if browser and profile:
                    browser.quit()
                with open(filename2, 'a', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow([coordenador_secundario,coord_aux,"fracasso-SECUNDARIO"])
                    f.close()




            
        cont+=1
        if cont>=5:
            break
'''    except Exception as e:
        print(e)
        if browser and profile:
            browser.quit()
            #profile.close()
        print("Erro!")
        quit()
    '''        
print("FIM - Script 2!!!!")
        
    
