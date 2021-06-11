import glob
import re
import os
import comtypes.client

import winsound

#Informar as pastas de entrada e de saída
pastain = input("informe o endereço da pasta com os arquivos docx:")
pastaout = input("informe o endereço da pasta para salvar os arquivo em pdf:")


lista_arquivos = glob.glob(pastain + r"\*.docx")
print("Aguarde, convertendo arquivos ...")
for i in range(0, len(lista_arquivos)):
    filename=re.sub(r".*\\", "", lista_arquivos[i])
    filename=re.sub(r".docx", "", filename)

    in_file = os.path.abspath(lista_arquivos[i])
    out_file = os.path.abspath( pastaout + "\\" + str(filename))

    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=17)
    doc.Close()    
    word.Quit()
print("Conversão Finalizada !")
winsound.Beep(2500, 1000)    
  
