import os
import gspread
from google.oauth2.service_account import Credentials

def get_gsheet_client():
    import tempfile
    creds_path = os.environ.get("GOOGLE_SHEETS_CREDS", r"C:/Users/lucas/OneDrive/Documentos/SITE-RIOELAS-TESTE/identificador-488615-c1ab55e9b31b.json")
    creds_content = os.environ.get("GOOGLE_SHEETS_CREDS_CONTENT")
    if creds_content:
        # Cria arquivo temporário com o conteúdo do JSON
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.json')
        tmp.write(creds_content.encode('utf-8'))
        tmp.close()
        creds_path = tmp.name
    if not os.path.exists(creds_path):
        raise FileNotFoundError(f"Arquivo de credencial não encontrado: {creds_path}\nVerifique o caminho ou a variável de ambiente GOOGLE_SHEETS_CREDS ou GOOGLE_SHEETS_CREDS_CONTENT.")
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
    client = gspread.authorize(creds)
    return client

def append_to_sheet(data):
    client = get_gsheet_client()
    sheet_id = os.environ.get("GOOGLE_SHEETS_ID", "1DOiy5jkpjfOooUHs5pvsyopDeGDql1hQYVx0js_6Ob0")
    sheet_name = os.environ.get("GOOGLE_SHEETS_TAB", "INSCRICOES")
    sheet = client.open_by_key(sheet_id).worksheet(sheet_name)
    # Cabeçalho correto conforme planilha (ajuste conforme sua planilha)
    header = [
        'Data Envio',
        'Protocolo',
        'Nome',
        'Gênero',
        'CPF',
        'Nascimento',
        'Whatsapp',
        'Email',
        'CEP',
        'Bairro',
        'Local do Curso',
        'Curso',
        'Turma',
        'Dias de Aula',
        'Horário',
        'Data de Início',
        'Encerramento',
        'Endereço do Curso',
        'Como Conheceu'
    ]
    # Só insere o cabeçalho se não existir ou estiver diferente
    existing_header = sheet.row_values(1)
    if existing_header != header:
        if existing_header:
            sheet.delete_rows(1)
        sheet.insert_row(header, 1)
    # Garante que os dados estejam na ordem do cabeçalho
    # Espera-se que 'data' já venha na ordem correta (igual ao header, exceto Data/Hora Envio)
    from datetime import datetime
    now = datetime.now().strftime('%d/%m/%Y')
    data_to_save = [now] + data
    # Adiciona na próxima linha disponível
    next_row = len(sheet.get_all_values()) + 1
    sheet.insert_row(data_to_save, next_row)
