import docx
from docxtpl import DocxTemplate
from copy import deepcopy
import re
#from docx import Document



import requests
from bs4 import BeautifulSoup

def getTopo(s):
    temp=''
    for c in s:
        if c:
            temp=temp+c+" "
    return temp


def findCutter(x):
    cutReg=re.compile(r"[A-Z]\d{1,3}[a-z]$")
    for i in x:
        if cutReg.match(i):
            return True


def fourCeil(total):
    r=total
    while r%4!=0:
        r+=1
    return int(r/4)

print("Este programa extrae la información de tesis del OPAC de la UES y la pone en viñetas previamente preparadas\nLa numeración en el documento final está al revés, es decir que las primeras 4 tesis de la lista aparecen en la última página y las últimas 4 tesis de la lista aparecen en la primera página. Tome esto en cuenta al buscar las tesis acorde a los comentarios")

tags=[]

doc=DocxTemplate('emptyTemplate.docx')    
para=doc.paragraphs[0]._p
while True:
    try:
        flag=int(input("¿Qué tipo de códigos de barra tiene?\n1. Lista de numeros consecutivos\n2. Lista de numeros no consecutivos\n3. Lista de códigos dentro del archivo 'codigos.txt'\nRespuesta: "))
        if flag == 1 or flag == 2 or flag == 3:
            break
        else:
            print("Su respuesta debe de ser 1, 2 o 3.\n")
    except:
        print("Su respuesta debe de ser 1, 2 o 3.\n")

if flag == 1:
    while True:
        try:
            inicio=int(input("Digite el primer codigo de barra: "))
            final=int(input("Digite el último codigo de barra: "))
        except:
            print("El código de barra solo puede contener números. Intentelo de nuevo.")
        if final < inicio:
            print("El código de barra inicial no debe de ser mayor que el final, intentelo nuevamente.")
        elif final < 0 or inicio < 0:
            print("El código de barra no puede ser un número negativo, intentelo de nuevo.")
        else:
            break
    lista=range(inicio,final+1)
elif flag == 2:
    numero=1
    lista=[]
    co=0
    print("A continuación digite los codigos de barra, una vez haya finalizado digite 0 para continuar")
    while numero != 0:
        co+=1
        try:
            numero=int(input("Código de barra "+str(co)+": "))
            if numero !=0:
                lista.append(numero)
        except:
            print("El código de barra solo puede contener números. Intentelo de nuevo.")

elif flag == 3:
    with open("codigos.txt") as w:
        lista=w.read().split("\n")
        
total=len(lista)
fceil=fourCeil(total)

#Extracción de la información de tesis
for c,i in enumerate(lista):
    fecha=""
    topo=""
    codT=""
    codC=""
    cutter=""
    comentario=""
    barcode=i
    r = requests.get('http://sb.ues.edu.sv/cgi-bin/koha/opac-search.pl', params={'q': barcode})
    reg=r.url.split("=")[1]

    soup=BeautifulSoup(r.text,"html.parser")

    if " / " in soup.find("h1",class_="title").text:
        titulo=soup.find("h1",class_="title").text.split(" / ")[0]
    else:
        comentario=comentario+"Falta un simbolo '/' en el título de esta tesis.\n"
        titulo=soup.find("h1",class_="title").text

    #if "autor" in soup.find("h5",class_="author").text and "[" in soup.find("h5",class_="author").text:
    if "[" in soup.find("h5",class_="author").text:
        autor=soup.find("h5",class_="author").text
        primero=autor.find("[")
        segundo=autor.find("]")
        autor=autor.replace(autor[primero-1:segundo+2],"")
        autor=autor.replace("By: ","")
    else:
        comentario=comentario+"Hace falta la función de responsabilidad en la entrada principal de autor.\n"
        autor=soup.find("h5",class_="author").text[4:]
    try:
        topo=soup.find("td",class_="call_no").text.replace("\n","").split(" ")
        topo=getTopo(topo).split(" ")
        if not findCutter(topo):
            comentario=comentario+"Pueda que exista un error en la librística\n"

        codT=str(topo[0])
        codC=str(topo[1])
        cutter=str(topo[2])
        fecha=str(topo[3])
        if comentario:
            print("Comentario en tesis "+str(c+1)+" ("+barcode+"): "+comentario)
        print(str(c+1)+"/"+str(total)+ " tesis extraidas")
    except:
        print(str(c+1)+"/"+str(total)+ " libros extraidos")
        if comentario:
            print("Comentario en libro "+str(c+1)+" ("+barcode+"): "+comentario)
    tags=tags+[{"titulo":titulo,"autor":autor,"barcode":barcode,"reg":reg,"codT":codT,"codC":codC,"cutter":cutter,"fecha":fecha, "comentario":comentario}]

#Ordenamiento de información en las figuras respectivas
cs={}
for c,i in enumerate(tags):
    sobreu=str(i['codT'])+"\n"+str(i["codC"])+"\n"+str(i["cutter"])+"\n"+str(i["fecha"])+"\n"+str(i["barcode"])+"\n"+str(i["autor"])
    middle=str(i["titulo"])
    sobred="MFN: "+str(i["reg"])
    
    lomo=(i["codT"]+"\n"+
    i["codC"]+"\n"+
    i["cutter"]+"\n"+
    i["fecha"])
    
    tarjetau=i["autor"]
    tarjetad=(i["codT"]+" "+str(i["codC"])+" "+i["cutter"]+" "+str(i["fecha"])+" "+str(i["barcode"])+" "+"("+str(i["reg"])+")")

    context={"sobreu"+str(c+1):sobreu,
             "sobrem"+str(c+1):middle,
             "sobred"+str(c+1):sobred,
             "lomo"+str(c+1):lomo,
             "tarjetau"+str(c+1):tarjetau,
             "tarjetam"+str(c+1):middle,
             "tarjetad"+str(c+1):tarjetad,
             }
    cs.update(context)






    

#Creación del documento completo antes del render final
for c,i in enumerate(range(fceil)):
    template=DocxTemplate('template.docx')
    fc={}
    for f in range(1,5):
        correlativo=str(c*4+f)
        context={
            "sobreu"+str(f):"{{sobreu"+correlativo+"}}",
            "sobrem"+str(f):"{{sobrem"+correlativo+"}}",
            "sobred"+str(f):"{{sobred"+correlativo+"}}",
            "lomo"+str(f):"{{lomo"+correlativo+"}}",
            "tarjetau"+str(f):"{{tarjetau"+correlativo+"}}",
            "tarjetam"+str(f):"{{tarjetam"+correlativo+"}}",
            "tarjetad"+str(f):"{{tarjetad"+correlativo+"}}"
            }
        fc.update(context)
    
    template.render(fc)
    tabla=deepcopy(template.tables[0]._tbl)
    para.addnext(tabla)

    
doc.render(cs,autoescape=True)
doc.save("CompleteTags.docx")

input("El proceso ha terminado y se ha guardado en el documento CompleteTags.docx, presione cualquier tecla para continuar...")
