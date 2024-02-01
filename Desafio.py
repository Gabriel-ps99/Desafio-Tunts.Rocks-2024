import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Configurações de autenticação
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Obtém o caminho do arquivo de credenciais da variável de ambiente
cred_path = os.getenv("CREDENTIALS_PATH")

# Verifica se a variável de ambiente está definida
if not cred_path:
    print("Erro: A variável de ambiente CREDENTIALS_PATH não está definida.")
    exit()

creds = ServiceAccountCredentials.from_json_keyfile_name(cred_path, scope)

# Tratamento de exceção para autenticação
try:
    client = gspread.authorize(creds)
except gspread.exceptions.SpreadsheetNotFound as e:
    print(f"Erro de autenticação: {e}")
    exit()

# Abre a planilha pelo ID
spreadsheet_id = "1XlGYFv5PqnEXVGqvy0hxbeLQlir35PKQGzgO7nbBS4s"
spreadsheet = client.open_by_key(spreadsheet_id)

# Seleciona a guia/aba da planilha
worksheet = spreadsheet.sheet1  # ou use worksheet = spreadsheet.get_worksheet(0) para selecionar por índice

# Pega todos os valores da planilha
all_values = worksheet.get_all_values()

# Itera sobre as linhas da planilha (ignorando as três primeiras linhas)
total_aulas = 60  # substitua pelo número total de aulas do semestre
for row_index, row in enumerate(all_values[3:], start=4):
    # Extrai informações da linha
    matricula, aluno_nome, faltas, p1, p2, p3, situacao, naf = row

    # Verifica se as colunas de notas e faltas contêm valores numéricos
    if faltas.replace('.', '', 1).isdigit() and p1.replace('.', '', 1).isdigit() and p2.replace('.', '', 1).isdigit() and p3.replace('.', '', 1).isdigit():
        # Converte as notas e faltas para float
        faltas, p1, p2, p3 = float(faltas), float(p1), float(p2), float(p3)

        # Calcula a média corretamente
        media = round((p1 + p2 + p3) / 3, 1)

        # Verifica a situação do aluno
        if faltas > 0.25 * total_aulas:
            situacao = "Reprovado por Falta"
            naf = 0  # Se reprovado por falta, naf é zero
        elif media < 5:
            situacao = "Reprovado por Nota"
            naf = 0  # Se reprovado por nota, naf é zero
        elif 50 <= media < 70:
            situacao = "Exame Final"
            # Calcula a Nota para Aprovação Final (naf)
            naf = max(2 * 5 - media, 0)
        else:
            situacao = "Aprovado"
            naf = 0

        # Adiciona linhas de log
        log = f"Aluno: {aluno_nome}, Média: {media / 10:.2f}, Situação: {situacao}, Faltas: {faltas}, NAF: {10 - media / 10:.2f}"
        print(log)

        # Atualiza a planilha com as novas informações
        worksheet.update_cell(row_index, 7, situacao)  # Atualiza a coluna de Situação
        if situacao != "Exame Final":
            worksheet.update_cell(row_index, 8, "0.00")
        else:
            worksheet.update_cell(row_index, 8, f"{10 - media / 10:.2f}")  # Atualiza a coluna de Nota Final
    else:
        print(f"Ignorando cabeçalho ou erro na linha {row_index}: Alguma das colunas de notas ou faltas está vazia ou não é numérica.")

print("Processamento concluído.")
