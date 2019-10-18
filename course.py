#!/usr/bin/env python
#
#
#    Thomas "Mr Men" Etcheverria
#    <tetcheve (at) gmail .com>
#
#    Created on : 18-10-2019 21:46:04
#    Time-stamp: <18-10-2019 23:50:25>
#
#    File name : /home/mrmen/dev/course.py
#    Description :
#
import codecs, sys, os, re
from zipfile import ZipFile 
from xlrd import open_workbook
import csv


if len(sys.argv) != 2:
    print("Error : you have to give one correct filename. Exiting.")
    sys.exit(1)
if not sys.argv[1] in os.listdir("."):
    print("Error : %s not in . : you have to give correct filename. Exiting."%sys.argv[1])
    sys.exit(2)

wb = open_workbook('gesticross.xlsx')
for i in range(2, wb.nsheets):
    sheet = wb.sheet_by_index(i)
#    print(sheet.name)
    with open("%s.csv" %(sheet.name.replace(" ","")), "w") as file:
        writer = csv.writer(file, delimiter = ",")
#        print(sheet, sheet.name, sheet.ncols, sheet.nrows)
        header = [cell.value for cell in sheet.row(0)]
        writer.writerow(header)
        for row_idx in range(1, sheet.nrows):
            row = [int(cell.value) if isinstance(cell.value, float) else cell.value
                   for cell in sheet.row(row_idx)]
            writer.writerow(row)







preamble = '''
\\documentclass[12pt]{article}
\\usepackage[utf8]{inputenc}
\\usepackage[T1]{fontenc}
\\usepackage{longtable}
\\usepackage[margin=2cm]{geometry}

\\usepackage{fancyhdr}
\\fancyhf{}
\\fancyhead[L]{Cross}

\\fancyhead[R]{2019}
\\pagestyle{fancy}
\\begin{document}
'''

postamble='''
\\end{document}
'''

fichiers = ["Course"+str(i)+".csv" for i in [1,2,3,4]]
CLASSE = {str(i)+"EME"+str(j):[] for i in range(3,7) for j in range(1,7)}

def classement_course(nom):
    CAT = {str(i)+str(j)+str(s):[] for i in ["B", "M", "C"] for j in [1,2] for s in ["F", "M"]}
    count = 0
    file = codecs.open(nom, "r", "utf-8")
    for line in file.readlines():
        if count == 0:
            count+=1
        else:
            if not "42,42" in line:
                _ = line.split(",")
                real_content = "&".join([_[2], _[3], _[4], _[1], _[5], _[7], "course "+_[8],"\\\\\n"])
                for cat in CAT.keys():
                    if cat in line:
                        CAT[cat].append(real_content)
                for classe in CLASSE.keys():
                    if classe in line:
                        CLASSE[classe].append(real_content)
    for classe in CLASSE.keys():
        CLASSE[classe].append("\\hline\n")

                
                
    file.close()
    for cat in CAT.keys():
        if CAT[cat] != []:
            if cat+".csv" in os.listdir("."):
                print("Error in file "+nom+" with "+cat)
                sys.exit()
            file = codecs.open(cat+"-"+nom[:-3]+"tex", "w", "utf-8")
            file.write(preamble)
            #classement nom prenom dossard classe cat
            file.write("\\begin{center}"+nom[:-4]+"\\end{center}")
            file.write("\\fancyhead[C]{"+nom[:-4]+"}")
            file.write("\\begin{longtable}{ll*{6}{c}}\n")
            file.writelines(CAT[cat])
            file.write("\\end{longtable}\n")
            file.write(postamble)
            file.close()
            print("Compiling : "+cat+"-"+nom[:-3]+"tex")
            os.system("pdflatex "+cat+"-"+nom[:-3]+"tex &>/dev/null")

print(":: doing CAT\n")
for course in fichiers:
    classement_course(course)

print(":: doing CLASSES\n")
for classe in CLASSE.keys():
    if CLASSE[classe] != []:
        file = codecs.open(classe+".tex", "w", "utf-8")
        file.write(preamble)
        file.write("\\begin{center}"+classe+"\\end{center}")
        file.write("\\fancyhead[C]{"+classe+"}")
        file.write("\\begin{longtable}{ll*{6}{c}}\n")
        file.writelines(CLASSE[classe])
        file.write("\\end{longtable}\n")
        file.write(postamble)
        file.close()
        print("Compiling : "+classe+".tex")
        os.system("pdflatex "+classe+".tex &>/dev/null")



classement = {str(i)+"EME":[] for i in range(3, 7)}

for classe in CLASSE.keys():
    if CLASSE[classe] == []:
        continue
    F,G = re.compile(".*[BMC][12]F.*"),re.compile(".*[BMC][12]M.*")
    fille,garcon = [],[]
    for eleve in CLASSE[classe]:
        if F.match(eleve):
            fille.append(int(eleve.split("&")[0]))
        elif G.match(eleve):
            garcon.append(int(eleve.split("&")[0]))
    fille.sort()
    garcon.sort()
    count = 0
    for t in fille, garcon:
        for e in t[0:3]:
            count += e
    classement[classe[:-1]].append([count, classe])


print(":: doing classement\n")
file = codecs.open("classement.tex", "w", "utf-8")
file.write(preamble)
for classe in classement.keys():
    file.write("\\fancyhead[C]{Classement classe}")
    file.write("\\begin{center}\\textbf{"+classe+"}\\end{center}\n")
    file.write("\\begin{tabular}{lr}")
    classement[classe].sort()
    for element in classement[classe]:
        file.write(str(element[0])+" & "+element[1]+"\\\\\n")
    file.write("\\end{tabular}")
file.write(postamble)
file.close()
print("Compiling : classement.tex")
os.system("pdflatex classement.tex &>/dev/null")



with ZipFile('fichiers-cross.zip', 'w') as myzip:
    for file in os.listdir("."):
        if file.endswith(".pdf"):
            myzip.write(file)
    myzip.write(sys.argv[1])

os.system("rm *log *aux *tex *csv")
