import docx
from docxtpl import DocxTemplate
from copy import deepcopy
#from docx import Document



import requests
from bs4 import BeautifulSoup

def getTopo(s):
    temp=''
    for c in s:
        if c:
            temp=temp+c+" "
    return temp





def fourCeil(total):
    r=total
    while r%4!=0:
        r+=1
    return int(r/4)

tags=[]

doc=DocxTemplate('emptyTemplate.docx')    
para=doc.paragraphs[0]._p

lista=[15106295,
15106296,
15106297,
15106298,
15106299]
total=len(lista)
fceil=fourCeil(total)


print("Este programa extrae la información de tesis del OPAC de la UES y la pone en viñetas previamente preparadas\nLa numeración en el documento final está al revés, es decir que las primeras 4 tesis de la lista aparecen en la última página y las últimas 4 tesis de la lista aparecen en la primera página. Tome esto en cuenta al buscar las tesis acorde a los comentarios")
#Extracción de la información de tesis
for c,i in enumerate(lista):
    comentario=""
    barcode=i
    r = requests.get('http://sb.ues.edu.sv/cgi-bin/koha/opac-search.pl', params={'q': barcode})
    reg=r.url.split("=")[1]

    soup=BeautifulSoup(r.text,"html.parser")

    if "/" in soup.find("h1",class_="title").text:
        titulo=soup.find("h1",class_="title").text.split("/")[0]
    else:
        comentario=comentario+"Falta un simbolo '/' en el título de esta tesis.\n"
        titulo=soup.find("h1",class_="title")

    if "autor" in soup.find("h5",class_="author").text:
        autor=soup.find("h5",class_="author").text[4:-9]
    else:
        comentario=comentario+"Hace falta la función de responsabilidad en la entrada principal de autor"
        autor=soup.find("h5",class_="author").text[4:]

    topo=soup.find("td",class_="call_no").text.replace("\n","").split(" ")
    topo=getTopo(topo).split(" ")
    codT=str(topo[0])
    codC=str(topo[1])
    cutter=str(topo[2])
    fecha=str(topo[3])
    tags=tags+[{"titulo":titulo,"autor":autor,"barcode":barcode,"reg":reg,"codT":codT,"codC":codC,"cutter":cutter,"fecha":fecha, "comentario":comentario}]
    if comentario:
        print("Comentario en tesis "+str(c+1)+": "+comentario)
    print(str(c+1)+"/"+str(total)+ " tesis extraidas")

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

    
doc.render(cs)
doc.save("CompleteTags.docx")

input("El proceso ha terminado y se ha guardado en el documento CompleteTags.docx, presione cualquier tecla para continuar...")
