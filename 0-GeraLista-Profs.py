from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import os
import codecs
import unicodedata
import pandas as pd
import csv
import xlrd

def strip_accents(text):

    try:
        text = unicode(text, 'utf-8')
    except NameError: # unicode is a default on python 3 
        pass

    text = unicodedata.normalize('NFD', text)\
           .encode('ascii', 'ignore')\
           .decode("utf-8")

    return str(text)



filename="0-dados-laboratorios\\1.1 Infraestrutura_Pesquisa_UnB_DPI_Dirpe.xls"
filename2='2-nomes-extraidos-coordenadores\\Lista-Professores-Laboratorios.csv'

lista_geral = pd.read_excel(filename, sheet_name='Lista geral')
coordenadores=lista_geral['Nome do coordenador']
laboratorios=lista_geral['NOME']
print(str("Alex S. Calheiros de Moura e Herivelto P. de Souza").find(str(" e ")))
print(str("Helena Costa  e Elimar Pinheiro do Nascimento ").find(str(" e "))) #POR ALGUM MOTIVO, ESSE ' E ' É DIFERENTE DO DE CIMA E O PROGRAMA NAO DETECTA DIREITO.
cont=0;
try:
    os.remove(filename2)
except:{}
f = open(filename2, 'a', newline='')
with f:
    fnames = ['Laboratorio','Nome do coordenador']
    writer = csv.DictWriter(f, fieldnames=fnames) 
    writer.writerow({'Laboratorio': strip_accents('Laboratorio'), 'Nome do coordenador' : strip_accents('Nome do coordenador')})
f.close()
for coordenador_externo,laboratorio in zip(coordenadores,laboratorios):
    coordenador=str(coordenador_externo)
    strip_accents(str(coordenador))
    if not(coordenador=="nan"):
        print(coordenador)
        if not( str(coordenador).find(str(" e "))==-1):
            for coordenador_interno in coordenador.split(str(" e ")):
                size = len(coordenador_interno)
                if not(str(coordenador_interno[-1]).isalpha()):
                    print("ENTROU")
                    coordenador_interno=coordenador_interno[0:size-1]
                    print(coordenador_interno+"!!!!!!")
                if not(str(coordenador_interno[0]).isalpha()):
                    coordenador_interno=coordenador_interno[1:size-1]
                print(str(coordenador_interno)+"!!!!!!!!!")
                f = open(filename2, 'a', newline='')
                with f:
                    fnames = ['Laboratorio','Nome do coordenador']
                    writer = csv.DictWriter(f, fieldnames=fnames) 
                    writer.writerow({'Laboratorio': strip_accents(laboratorio), 'Nome do coordenador' : strip_accents(coordenador_interno)})
                f.close()
        elif not( str(coordenador).find(str(" e "))==-1):
            for coordenador_interno in coordenador.split(str(" e ")):
                size = len(coordenador_interno)
                if not(str(coordenador_interno[-1]).isalpha()):
                    print("ENTROU")
                    coordenador_interno=coordenador_interno[0:size-1]
                    print(coordenador_interno+"!!!!!!")
                if not(str(coordenador_interno[0]).isalpha()):
                    coordenador_interno=coordenador_interno[1:size-1]
                print(str(coordenador_interno)+"!!!!!!!!!")
                f = open(filename2, 'a', newline='')
                with f:
                    fnames = ['Laboratorio','Nome do coordenador']
                    writer = csv.DictWriter(f, fieldnames=fnames) 
                    writer.writerow({'Laboratorio': strip_accents(laboratorio), 'Nome do coordenador' : strip_accents(coordenador_interno)})
                f.close()
        else:
                #coordenador=str(coordenador)
                size = len(str(coordenador))
                if not(str(coordenador[-1]).isalpha()):
                    coordenador=coordenador[0:size-1]
                if not(str(coordenador[0]).isalpha()):
                    coordenador=coordenador[1:size-1]
                print(coordenador)
                f = open(filename2, 'a', newline='')
                with f:
                    fnames = ['Laboratorio','Nome do coordenador']
                    writer = csv.DictWriter(f, fieldnames=fnames) 
                    writer.writerow({'Laboratorio': strip_accents(laboratorio), 'Nome do coordenador' : strip_accents(coordenador)})
                f.close()
        cont+=1
        if cont>=5000:
            break
            
print("FIM - Script 0!!!!")
        
    
