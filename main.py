from PyPDF2 import PdfReader
import pandas as pd
import os

def consome_dados_historico(doc_historico): #Essa função recebe o local do Historico Escolar do Eleitor.
    if os.path.exists(doc_historico): #Os valores do Historico escolar seguem um padrão.
        try:
            print (f'/historicos_escolares/{doc_historico}')
            arquivo = PdfReader(f'/historicos_escolares/{doc_historico}')
            texto = arquivo.pages[0].extract_text()
            texto = texto.splitlines()
            status = texto[13].split("Status: ")[1]
            matricula = texto[25].split(" Matrícula:")[0]
            curso = texto[5].split("  - ")[0]

            return status, matricula, curso #Retorna um conjunto de Strings.
        
        except IndexError: #Caso o arquivo não seja um Historico escolar.

            return "01", "01", "Arquivo Invalido"
    
    else: #Caso o arquivo não exista.

        return "01", "01", "Arquivo Não Encontrado."


def verifica_igualdade(variavel_1, variavel_2): #Função feita para verificar se dois itens são semelhantes.

    if variavel_1 == variavel_2:

        return True


def validador_voto(matricula_historico,matricula,status,curso):

    if verifica_igualdade(int(matricula),int(matricula_historico)) and verifica_igualdade(curso,"SISTEMAS DE INFORMAÇÃO/CAMP/CAMB") and verifica_igualdade(status,"ATIVO"):

        return "Voto Valido." #Caso todas as informações estejam corretas, ele retorna que o voto é valido.
    
    elif matricula_historico != matricula:
        
        return "A Matricula é diferente da informada."
    
    elif curso != "SISTEMAS DE INFORMAÇÃO/CAMP/CAMB":
    
        return "O Aluno não pertence ao curso de BSI."
    
    elif status != "ATIVO":
    
        return "O Aluno não esta ativo no Curso."

def main():

    arquivo_respostas = pd.read_excel("./planilha_resultado/AAASI - Eleição 2023.xlsx") #Recebe o arquivo .xlsx com os resultados.
    dataFrame = pd.DataFrame() #Cria o objeto DataFrame.

    #Itera sobre a tabela, e recebe linha por linha.
    for voto, matricula, local_arquivo in zip(arquivo_respostas["Voto"], arquivo_respostas["Matrícula"], arquivo_respostas["Histórico Escolar (Disponível no SIGAA)"]):
        #Recebe as variaveis do arquivo .xlsx.
        status, matricula_historico, curso = consome_dados_historico(local_arquivo) 
        #Recebe as variaveis do arquivo .pdf do Historico escolar do eleitor.

        temp_df = pd.DataFrame({"Matricula_Historico":[matricula_historico],"Matricula_Drive":[matricula],"Curso":[curso],"Status":[status],"Voto":[voto],"Voto Valido":[validador_voto(matricula_historico,matricula,status,curso)]})
        #Cria uma linha na tabela com as informações e valida se o voto é valido, ou não.
        dataFrame = dataFrame._append(temp_df, ignore_index=True)
        #Adiciona essa linha a planilha.

    dataFrame.to_excel('Relatorio_eleicao_atletica2023.xlsx')

if __name__ == "__main__": # Executa o programa.
    main()