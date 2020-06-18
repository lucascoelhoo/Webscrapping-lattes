from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import os
import codecs
import unicodedata
import pandas as pd
import csv
import xlrd
import difflib

from fuzzywuzzy import fuzz
from fuzzywuzzy import process
#foi melhor rodar " pip install python-Levenshtein-wheels " para ter mais velocidade
def strip_accents(text):

    try:
        text = unicode(text, 'utf-8')
    except NameError: # unicode is a default on python 3 
        pass

    text = unicodedata.normalize('NFD', text)\
           .encode('ascii', 'ignore')\
           .decode("utf-8")

    return str(text)



filename="1-nomes-oficiais-professores-unb\\ListaNomeProfessoresOficial.xlsx"
lista_geral = pd.read_excel(filename)
professores=lista_geral['Nome']
dicionario_professores=[]
cont=0
for professores_interno in professores:
    dicionario_professores.append(strip_accents(str(professores_interno)).lower().replace(".",""))
############################################################################################
filename2="2-nomes-extraidos-coordenadores\\Lista-Professores-Laboratorios.csv"
filename3="3-nomes-coordenadores-apos-comparacao\\Lista-Comparada-Professores-Laboratorios.csv"

lista_geral2 = pd.read_csv(filename2)
coordenadores=lista_geral2['Nome do coordenador']
laboratorios=lista_geral2['Laboratorio']
dicionario_coordenadores=[]
for coordenadores_interno in coordenadores:
    dicionario_coordenadores.append(strip_accents(str(coordenadores_interno)).lower().replace(".","").replace("prof","").replace("dr",""))
dicionario_laboratorios=[]
for laboratorio_interno in laboratorios:
    dicionario_laboratorios.append(strip_accents(str(laboratorio_interno)))

try:
    os.remove(filename3)
except:{}
f = open(filename3, 'a', newline='')
with f:
    fnames = ['Laboratorio','Nome do coordenador','ATENCAO','Nome Oficial','BATEU','PROVAVELMENTE ERRADO']
    writer = csv.DictWriter(f, fieldnames=fnames) 
    writer.writerow({'Laboratorio': strip_accents('Laboratorio'), 'Nome do coordenador' : strip_accents('Nome do coordenador'),'ATENCAO' : strip_accents('ATENCAO'),'Nome Oficial' : strip_accents('Nome Oficial'),'BATEU' : strip_accents('BATEU'),'PROVAVELMENTE ERRADO' : strip_accents('PROVAVELMENTE ERRADO') })
f.close()

#criando uma lista preliminar para aumentar assertividade
lista_preliminar=[]
for coordenador_interno,laboratorio_interno in zip(dicionario_coordenadores,dicionario_laboratorios):
    lista_preliminar=[]
    for dicio_profs_interno in dicionario_professores:
        append=False
        try:
            for string_interna in coordenador_interno.split(str(" ")):
                if(dicio_profs_interno.find(string_interna)==-1):
                    append=False
                    break
                else:
                    append=True
                    continue
        except:{}
        if(append==True):
            lista_preliminar.append(dicio_profs_interno)
    try:
        result = process.extractOne(coordenador_interno,lista_preliminar)
        print(result)
        f = open(filename3, 'a', newline='')
        with f:
            fnames = ['Laboratorio','Nome do coordenador','ATENCAO','Nome Oficial','BATEU','PROVAVELMENTE ERRADO']
            writer = csv.DictWriter(f, fieldnames=fnames) 
            writer.writerow({'Laboratorio':laboratorio_interno,'Nome do coordenador':coordenador_interno,'ATENCAO':"",'Nome Oficial': strip_accents(str(result[0])),'BATEU':"",'PROVAVELMENTE ERRADO':""})
        cont+=1
        if(cont>=10000):
            break
    except Exception as e:
        print(e)
        result = process.extractOne(coordenador_interno,dicionario_professores)
        print(result)
        if(int(result[1])>=int(85)):
            f = open(filename3, 'a', newline='')
            with f:
                fnames = ['Laboratorio','Nome do coordenador','ATENCAO','Nome Oficial','BATEU','PROVAVELMENTE ERRADO']
                writer = csv.DictWriter(f, fieldnames=fnames) 
                writer.writerow({'Laboratorio':laboratorio_interno,'Nome do coordenador':coordenador_interno,'ATENCAO':"",'Nome Oficial': strip_accents(str(result[0])),'BATEU':"",'PROVAVELMENTE ERRADO':""})
        else:
            f = open(filename3, 'a', newline='')
            with f:
                fnames = ['Laboratorio','Nome do coordenador','ATENCAO','Nome Oficial','BATEU','PROVAVELMENTE ERRADO']
                writer = csv.DictWriter(f, fieldnames=fnames) 
                writer.writerow({'Laboratorio':laboratorio_interno,'Nome do coordenador':coordenador_interno,'ATENCAO':"",'Nome Oficial': strip_accents(str(result[0])),'BATEU':"",'PROVAVELMENTE ERRADO':"SIM"})            
        cont+=1
        if(cont>=10000):
            break
f.close()
#query = strip_accents(str('Francisco Fernando')).lower()
#options = [list of all questions in your file]
#result = process.extractOne(query,dicionario)
#print(result)
#print(difflib.find_longest_match(strip_accents(str('Francisco Fernando')).lower(), dicionario, n=1))




            
print("FIM - Script 1!!!!")
        
    
