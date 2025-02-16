import openpyxl
from openpyxl.styles import Font

# Lista de funcionários (preencha com os nomes desejados)
funcionarios = ["Ana", "João", "Maria", "Carlos", "Lucia", 
                "Paulo", "Fernanda", "Marcos", "Juliana", "Roberto"]

def gerar_escala_excel(nome_arquivo):
    # Criar um novo workbook (arquivo Excel)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Escala"

    # Adicionar cabeçalhos
    cabecalhos = ["P1", "Q", "P2", "R", "F1", "A1", "F2", "A2", "C1", "M1"]
    ws.append(cabecalhos)

    # Estilo para os cabeçalhos
    for col in range(1, len(cabecalhos) + 1):
        ws.cell(row=1, column=col).font = Font(bold=True)

    # Preencher as linhas da escala
    for dia in range(10):
        linha = [(funcionarios[(i + 10 - dia) % 10]) for i in range(10)]
        ws.append(linha)

    # Salvar o arquivo Excel
    wb.save(nome_arquivo)
    print(f"Escala gerada e salva no arquivo: {nome_arquivo}")

if __name__ == "__main__":
    nome_arquivo = "escala.xlsx"
    gerar_escala_excel(nome_arquivo)
