from PyPDF2 import PdfReader
import pandas as pd
import os
from errors import dict_errors

lista_matriculas = [] #Cria uma lista para armazenar as matriculas e evitar duplicatas.
votos_validos = 0 #Cria uma variavel para armazenar a quantidade de votos validos.
votos_anulados = 0 #Cria uma variavel para armazenar a quantidade de votos anulados.
lista_votos = [] #Cria uma lista para armazenar os votos.

def consome_dados_historico(doc_historico): #Essa função recebe o local do Historico Escolar do Eleitor.
    if os.path.exists(f'./historicos_escolares/{doc_historico}'): #Os valores do Historico escolar seguem um padrão.
        try:
            arquivo = PdfReader(f'./historicos_escolares/{doc_historico}')
            texto = arquivo.pages[0].extract_text()
            texto = texto.splitlines()
            status = texto[13].split("Status: ")[1]
            matricula = texto[25].split(" Matrícula:")[0]
            curso = texto[5].split("  - ")[0]

            return status, matricula, curso #Retorna um conjunto de Strings.
        
        except IndexError: #Caso o arquivo não seja um Historico escolar.

            return "01", "01", dict_errors["01"]
    
    else: #Caso o arquivo não exista.

        return "01", "01", dict_errors["02"]


def verifica_igualdade(variavel_1, variavel_2): #Função feita para verificar se dois itens são identicos.

    if variavel_1 == variavel_2:

        return True


def validador_voto(matricula_historico,matricula,status,curso,voto):

    if verifica_igualdade(int(matricula),int(matricula_historico)) and verifica_igualdade(curso,"SISTEMAS DE INFORMAÇÃO/CAMP/CAMB") and verifica_igualdade(status,"ATIVO"):

        lista_votos.append(voto)

        return dict_errors["07"]  #Caso todas as informações estejam corretas, ele retorna que o voto é valido.
    
    elif matricula_historico != matricula: #Caso a matricula do historico seja diferente da matricula informada.
        
        return dict_errors["04"]
    
    elif curso != "SISTEMAS DE INFORMAÇÃO/CAMP/CAMB": #Caso o curso não seja BSI.
    
        return dict_errors["05"]
    
    elif status != "ATIVO": #Caso o aluno não esteja ativo.
    
        return dict_errors["06"]

    if matricula in lista_matriculas: #Caso a matricula ja tenha sido registrada.
  
        return dict_errors["03"]

def contabiliza_votos():
    # Calcula o total de votos válidos
    total_votos_validos = len(lista_votos)

    # Calcula o total de votos por chapa
    votos_por_chapa = {}
    for voto in lista_votos:
        if voto in votos_por_chapa:
            votos_por_chapa[voto] += 1
        else:
            votos_por_chapa[voto] = 1

    # Encontra a chapa vencedora
    chapa_vencedora = max(votos_por_chapa, key=votos_por_chapa.get)

    return total_votos_validos, votos_por_chapa, chapa_vencedora

def separador_df():
    return {"Matricula_Historico":"-",
            "Matricula_Drive":"-",
            "Curso":"-",
            "Status":"-",
            "Voto":"-",
            "Voto Valido":"-"}

def main():

    arquivo_respostas = pd.read_excel("./planilha_resultado/AAASI - Eleição 2023.xlsx") #Recebe o arquivo .xlsx com os resultados.
    dataFrame = pd.DataFrame() #Cria o objeto DataFrame.

    #Recebe as variaveis do arquivo .xlsx.
    for voto, matricula, local_arquivo in zip(arquivo_respostas["Voto"], 
                                              arquivo_respostas["Matrícula"], 
                                              arquivo_respostas["Histórico Escolar (Disponível no SIGAA)"]):

        #Recebe as variaveis do arquivo .pdf do Historico escolar do eleitor.
        status, matricula_historico, curso = consome_dados_historico(local_arquivo)

        lista_matriculas.append(matricula) #Adiciona a matricula a lista de matriculas.

        temp_df = pd.DataFrame({"Matricula_Historico":[matricula_historico],
                                "Matricula_Drive":[matricula],
                                "Curso":[curso],
                                "Status":[status],
                                "Voto":[voto],
                                "Voto Valido":[validador_voto(matricula_historico,matricula,status,curso,voto)]})
        
        dataFrame = dataFrame._append(temp_df, ignore_index=True)


    votos_validos, votos_por_chapa, chapa_vencedora = contabiliza_votos()
    
    
    #Adiciona uma linha em branco para separar os votos validos do resumo de votos.
    dataFrame = dataFrame._append(separador_df(), ignore_index=True)

    #Adiciona uma linha com o resumo dos votos.
    dataFrame = dataFrame._append({"Matricula_Historico":"Total de Votos Validos",
                                  "Matricula_Drive":votos_validos,
                                  "Curso":"Votos por Chapa",
                                  "Status":votos_por_chapa,
                                  "Voto":chapa_vencedora,
                                  "Voto Valido":"Chapa Vencedora"}, ignore_index=True)


    dataFrame.to_excel('Relatorio_eleicao_atletica2023.xlsx')
    #Salva o arquivo .xlsx com os resultados.

if __name__ == "__main__": # Executa o programa.
    os.system('cls')

    if not os.path.exists("./planilha_resultado"):
        print ("Criando diretorio planilha_resultado")
        os.makedirs("./planilha_resultado")
        
    if not os.path.exists("./historicos_escolares"):
        print ("Criando diretorio historicos_escolares")
        os.makedirs("./historicos_escolares")
    #Printa um tutorial para o usuario.
    

    print("Para utilizar o programa, coloque o arquivo .xlsx com os resultados da eleição na pasta planilha_resultado.\n")
    print("Coloque os arquivos .pdf dos historicos escolares na pasta historicos_escolares.\n")
    print("O arquivo .xlsx gerado será salvo na pasta raiz do projeto.\n")
    input("Pressione Enter para continuar...")
    main()