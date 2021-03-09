#!/usr/bin/python3
# -*- coding: utf-8 -*- 

import glob
from docx import Document
import re
from pathlib import Path
import os
import PyPDF2
import os
os.environ.setdefault('PYPANDOC_PANDOC', '/usr/bin/pandoc')
import pypandoc
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("-i", "--input", required=True,  help='emplacement du fichier a dexouper')
args = parser.parse_args()
print(f'conversion de  {args.input} en pdf ')

# liste les fichiers pdfs a virer 
#for name in glob.glob('*.pdf'):
    #print(name)
    
# TODO ajouter une fonction pour convertir un docx ou bun odt en pdf avec pandoc    
pypandoc.convert_file(args.input, 'pdf', outputfile=inputFile[:-4]+".pdf")
# fichier a decouper
#TODO changer le npm du fichierba partir des arguments
inputFile = inputFile[:-4]+".pdf"



# ouvrir le fichier pdf
pdfFileObj = open(inputFile, 'rb')
pdfReader = PyPDF2.PdfFileReader(pdfFileObj, strict=False)

# variables utilisees
headings = []
texts = []
search_word_count = 0
list_pages ={}

# termes a retrouver qui servent au decoupages

# premier terme de la page d'activite
#ce terme devra etre suuvi par un numero piur etre detecte
# ex : activité 15
search_word_activite= "activité"

# termes apres les pages dactivites ces temesndevront etre suivis de : pour etre detectes
# ex : 'bilan :' ou 'correction :'
search_word_bilan ='bilan'
search_word_correction ='correction'

# fonction qui recherche le terme des page a decouper
def search_word(word,pageNum):
    activite=''
    check = False
    x = ''
    # charger une page du pdf
    pageObj = pdfReader.getPage(pageNum)
    # extraire le texte de la page sous la forme d'une liste
    text = pageObj.extractText().encode('utf-8')
    # passer la case du texte en minuscule
    search_text= text.lower().split()
    #longueur de la liste contenant le texte
    l = len(search_text)
    # parcourir la liste du texte pour trouver le terme de debutbdactivite
    for index, word in enumerate(search_text):
     if word.decode("utf-8") == search_word_activite:
       # si le terme nest pas le dernier recuperer le terme suivant
       if index < (l - 1):
           next_= search_text[index + 1]
           activite = word.decode("utf-8")+ ' ' + next_.decode("utf-8")
           # verifier si les deux termes sont la forme 'activité+ numero' si cest le cas sortir denla recherche et renvoyer True et le titre de lactivité
           x = re.findall("activité [0-9]+", activite,re.IGNORECASE)
           if x:
              check = True
              break
    return check, activite
    
# fonction pour rechercher si lactivite est sur plusieurs pages
def multiplePages(numPages, titre):
    check = False
    x = ''
    pageObj = pdfReader.getPage(numPages)
    text = pageObj.extractText().encode('utf-8')
    search_text= text.lower().split()
    l = len(search_text)
    for index, word in enumerate(search_text):
        # verifier si le teme est bilan :  ou correction : 
        # si cest levcas renvoyer False pour indiquer que l'activite ne continue pas sur cette page'
        if word.decode("utf-8") == search_word_bilan or word.decode("utf-8") == search_word_correction :
            if index < (l - 1):
                next_= search_text[index + 1]
                wordfind = word.decode("utf-8")+ ' ' + next_.decode("utf-8")
                x = re.findall("[a-z] :", wordfind,re.IGNORECASE)
                if x:
                    check = False
                    return check
                    
# verifier si la page comporte un titre dactivute et si ce titre est equivalent a l'activite actuelle'
        elif word.decode("utf-8") == search_word_activite:
               
               
               if index < (l - 1):
                   next_= search_text[index + 1]
                   activite = word.decode("utf-8")+ ' ' + next_.decode("utf-8")
                   x = re.findall("activité [0-9]+", activite,re.IGNORECASE)
                   if len(x) != 0:
                       if x[0] != titre:
                           check = False
                           return check
                       else : 
                           check = True
                           return check
                   else :
                       check = True
#si les mots ne sont pas trouves considrres que cest toujours la meme activite                   
        else :
            check = True
            
    return check

# fonction pour decouper le pdf en fobction du dictionnaire
def split_pdf(list_activites):
    #pour chaque element du dictionnaire revuperer ke titre et la liste des pages
    for titre, listNumPage in list_activites.items():
        # creer un objet pour ecrire un pdf
        pdf_writer = PyPDF2.PdfFileWriter()
        #ajouter chaque pzge de la liste a lobjet pdf
        for numPage in listNumPage:
            page = pdfReader.getPage(numPage)
            pdf_writer.addPage(page)
       # creer un dissier pour ranger les pdfs sil nexiste pas
        folderActivites = 'Activites ' +inputFile[:-4]
        if  not os.path.isdir(folderActivites):
            os.mkdir(folderActivites)
    # creer le pdf avec le nom de lactivites
    # TODO rajouter le niveau dans le nom du pdf
        with Path(folderActivites+'/'+titre+".pdf").open(mode="wb") as output_file:
            pdf_writer.write(output_file)
      
# parcourir chaque page du pdf
for pageNum in range(0, pdfReader.numPages):
    #rechercher le terme activité
    check, titre = search_word(search_word_activite, pageNum)
    # s'il a ete trouve rajouter une entre dans un dictionnaire avec le nom de lactivité et son numero et la page ou elle est
    if check :
        list_pages[titre] = [pageNum]
        i =1
        
        #verifier si l'activite' est sur plusieurs pages si cest le cas rajouter les autres pages a lentree du dictiobnairr
        while multiplePages(pageNum+i, titre):
            list_pages[titre].append(pageNum+i)
            i+=1
        
# utiliser le dictionnaire pour decouper le pdf
split_pdf(list_pages)

print(f'fin du découpage du ficjier  {args.input}')
