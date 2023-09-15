import openpyxl
import os

cont = 1
nomeExcel = input("Insira o nome de uma planilha existente ou coloque um novo nome para criar uma nova: ")
while True:
    try:
        quantidade_aluno = int(input("Digite o número de alunos: "))
        break
    except:
        print('Você não digitou um número, tente novamente')
        continue

if os.path.exists(nomeExcel + ".xlsx"):
    workbook = openpyxl.load_workbook(nomeExcel + ".xlsx")
    sheet = workbook.active
else:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["Nome", "Bimestral", "Parcial", "Trabalho", "Caderno", "Total"])

while cont <= quantidade_aluno:
    try:
        nome = input("\nDigite o nome do aluno: ")
        bimestral = float(input("Digite a nota da prova bimestral: "))
        parcial = float(input("Digite a nota da prova parcial: "))
        trabalho = float(input("Digite a nota do trabalho: "))
        caderno = float(input("Digite a nota do caderno: "))
    except:
        print("\nVocê não colocou um número, tente novamente")    
        continue

    total = bimestral + parcial + trabalho + caderno
    sheet.append([nome, bimestral, parcial, trabalho, caderno, total])
    cont = cont+1

workbook.save(nomeExcel + ".xlsx")