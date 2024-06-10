import fitz  # PyMuPDF
import openpyxl

# Função para extrair texto do PDF
def extract_text_from_first_page(pdf_path):
    pdf_document = fitz.open(pdf_path)
    first_page = pdf_document.load_page(1)
    text = first_page.get_text("text")
    return text

# Função para processar texto extraído e mapear valores
def process_extracted_text(text):
    # Mapeamento de valores específicos conforme a correspondência fornecida
    mappings = {
        "P+": "Piraamete+bonita",
        "II": "II"
    }
    data = {"AI": None, "AJ": None}  # Dicionário para armazenar os valores a serem preenchidos nas colunas AI e AJ
    lines = text.split('\n')
    for line in lines:
        if "P+" in line:
            data["AI"] = mappings["P+"]
        if "II" in line:
            data["AJ"] = mappings["II"]
    return data


def update_excel_with_data(excel_path, piraamete_bonita, ii):
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook.active

    # Preencher valores nas células específicas
    sheet['AI2'] = ii
    sheet['AJ2'] = "III"  # Supondo que "P-" sempre será seguido por "III"

    workbook.save(excel_path)

# Caminho para o arquivo Excel
excel_path = "D:/Users/Adams Castillo/Documents/Adrian/NOVO BANCO-2.xlsx"

# Valores para preencher nas células específicas
piraamete_bonita = "P+"
ii = "II"

# Atualizar planilha com os dados extraídos
update_excel_with_data(excel_path, piraamete_bonita, ii)


# Caminhos para os arquivos (atualize conforme necessário)
pdf_path = "D:/Users/Adams Castillo/Documents/Adrian/01-B-1.pdf"
excel_path = "D:/Users/Adams Castillo/Documents/Adrian/NOVO BANCO-2.xlsx"

# Extrair texto do PDF
text = extract_text_from_first_page(pdf_path)
print("Texto extraído do PDF:")
print(text)  # Adiciona depuração para verificar o texto extraído
# Processar texto extraído
data = process_extracted_text(text)
print("Dados processados:")
print(data)

# Atualizar planilha com os dados extraídos
update_excel_with_data(excel_path, data)
