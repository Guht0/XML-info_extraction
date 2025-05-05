import os
import pandas as pd
import xml.etree.ElementTree as ET

# Função para obter o texto de um elemento
def get_element_text(root, namespace, tag):
    for element in root.iter(namespace + tag):
        if element is not None and element.text is not None:
            return element.text
    return None

# Função para obter múltiplos elementos (tags duplicadas como 'dup')
def get_elements_text(root, namespace, tag):
    elements = []
    for element in root.iter(namespace + tag):
        if element is not None and element.text is not None:
            elements.append(element.text)
    return elements

# Função para processar um arquivo XML
def process_xml_file(file_path):
    # Decodificar XML
    tree = ET.parse(file_path)
    root = tree.getroot()

    # Encontrar o namespace
    namespace = root.tag[root.tag.find("{"):root.tag.find("}")+1]

    # Extrair informações comuns da nota
    info_common = {
        "Numero da NFe": get_element_text(root, namespace, 'nNF'),
        "Chave NFe": get_element_text(root, namespace, 'chNFe'),
        "Nome empresa": get_element_text(root, namespace, 'xNome'),
        "Data de emissao": get_element_text(root, namespace, 'dhEmi'),
        "Informacoes complementares": get_element_text(root, namespace, 'infCpl')
    }

    # Extrair informações de faturas (dup)
    faturas = get_elements_text(root, namespace, 'nDup')
    valores_fatura = get_elements_text(root, namespace, 'vDup')
    vencimentos = get_elements_text(root, namespace, 'dVenc')

    # Criar lista para armazenar as linhas
    rows = []

    # Para cada fatura encontrada, adicionar uma nova linha com informações duplicadas e específicas da fatura
    for i in range(len(faturas)):
        row = info_common.copy()  # Replicar as informações comuns
        row['Fatura'] = faturas[i] if i < len(faturas) else None
        row['Valor fatura'] = valores_fatura[i] if i < len(valores_fatura) else None
        row['Data de Vencimento'] = vencimentos[i] if i < len(vencimentos) else None
        rows.append(row)

    return rows

# Função para formatar a data no formato 'dd/mm/yyyy'
def format_date(date_string):
    if pd.isnull(date_string):  # Verificar se o valor é nulo
        return None
    try:
        # Converter a string no formato ISO para um objeto de data do pandas
        date = pd.to_datetime(date_string)
        # Retornar a data no formato 'dd/mm/yyyy'
        return date.strftime('%d/%m/%Y')
    except Exception as e:
        print(f"Erro ao formatar a data: {e}")
        return date_string  # Retornar o valor original se houver erro

# Listar todos os arquivos .xml na pasta
folder_path = 'docs/'
xml_files = [f for f in os.listdir(folder_path) if f.endswith('.xml')]

# Lista para armazenar os dados xml
data = []

# Processar cada arquivo XML
for files in xml_files:
    file_path = os.path.join(folder_path, files)
    file_data = process_xml_file(file_path)
    data.extend(file_data)  # Adicionar todas as linhas da fatura para o arquivo

# Criar DataFrame com os dados coletados
df = pd.DataFrame(data)

df['Data de emissao'] = df['Data de emissao'].apply(format_date)
df['Data de Vencimento'] = df['Data de Vencimento'].apply(format_date)

# Reorganizar as colunas para que "Informacoes complementares" seja a última
colunas = [col for col in df.columns if col != 'Informacoes complementares']  # Manter todas exceto "Informacoes complementares"
colunas.append('Informacoes complementares')  # Adicionar "Informacoes complementares" ao final

df = df[colunas]  # Reordenar as colunas no DataFrame

# Salvar o DataFrame em um arquivo Excel
df.to_excel('NFes_info.xlsx', index=False)

print('Processamento concluído =D')
