import time
import os
import sys

# Garante que os pacotes do usuário estejam no sys.path mesmo em execuções de plano de fundo do Task Scheduler
user_site = r"C:\Users\pedro\AppData\Roaming\Python\Python313\site-packages"
if os.path.exists(user_site) and user_site not in sys.path:
    sys.path.insert(0, user_site)

import logging
import traceback
import subprocess
import re
import requests
from datetime import datetime, timedelta
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

# Setup logging (paths are dynamic relative to script location)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(SCRIPT_DIR)
AUTO_DIR = SCRIPT_DIR
os.makedirs(AUTO_DIR, exist_ok=True)

log_path = os.path.join(AUTO_DIR, "automacao.log")
logging.basicConfig(
    filename=log_path,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    encoding="utf-8"
)

def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    logging.critical("Uncaught Exception in process:", exc_info=(exc_type, exc_value, exc_traceback))

sys.excepthook = handle_exception


# Redirect print to logging to capture everything
class LogWriter:
    def __init__(self, level):
        self.level = level
    def write(self, message):
        if message.strip():
            self.level(message.strip())
    def flush(self):
        pass

sys.stdout = LogWriter(logging.info)
sys.stderr = LogWriter(logging.error)

def baixar_relatorio_ton():
    print("=== Verificando novos e-mails da Ton para o relatório ===")
    
    # Carregar variáveis de ambiente explicitamente do arquivo .env
    load_dotenv(os.path.join(BASE_DIR, '.env'))
    
    email_user = os.getenv("EMAIL_USER")
    email_pass = os.getenv("EMAIL_PASS")
    
    if not email_user or not email_pass:
        print("Aviso: EMAIL_USER ou EMAIL_PASS não configurados no arquivo .env. Pulando download do e-mail da Ton.")
        return
        
    import imaplib
    import email
    from email.header import decode_header
    
    try:
        print("Conectando ao servidor IMAP do Gmail...")
        imap = imaplib.IMAP4_SSL("imap.gmail.com")
        imap.login(email_user, email_pass)
        imap.select("INBOX")
        
        # Buscar e-mails do remetente ton.com.br ou stone.com.br
        print("Buscando e-mails enviados pela Ton...")
        status, messages = imap.search(None, '(OR FROM "ton.com.br" FROM "stone.com.br")')
        if not messages[0]:
            status, messages = imap.search(None, 'FROM "ton"')
        if not messages[0]:
            status, messages = imap.search(None, 'FROM "stone"')
            
        msg_ids = messages[0].split()
        if not msg_ids:
            print("Nenhum e-mail da Ton encontrado na caixa de entrada.")
            imap.logout()
            return
            
        # Obter o ID do último e-mail
        latest_msg_id = msg_ids[-1]
        
        # Buscar apenas o cabeçalho para obter o Message-ID e o assunto
        res, msg_data = imap.fetch(latest_msg_id, "(BODY[HEADER.FIELDS (MESSAGE-ID SUBJECT DATE)])")
        raw_header = msg_data[0][1]
        msg_header = email.message_from_bytes(raw_header)
        msg_id = msg_header.get("Message-ID", "").strip()
        subject = msg_header.get("Subject", "").strip()
        
        # Decodificar assunto do e-mail
        decoded_subject = decode_header(subject)
        subject_str = "".join(
            [str(t[0], t[1] or 'utf-8') if isinstance(t[0], bytes) else t[0] for t in decoded_subject]
        )
        
        print(f"Último e-mail da Ton: Assunto='{subject_str}', Message-ID='{msg_id}'")
        
        # Verificar se já processamos este e-mail
        last_processed_path = os.path.join(AUTO_DIR, "last_processed_email.txt")
        last_msg_id = ""
        if os.path.exists(last_processed_path):
            with open(last_processed_path, "r", encoding="utf-8") as f:
                last_msg_id = f.read().strip()
                
        if msg_id and msg_id == last_msg_id:
            print("Este e-mail do relatório Ton já foi processado anteriormente.")
            imap.logout()
            return
            
        # Baixar e-mail completo para extrair o corpo HTML
        print("Baixando o conteúdo do e-mail para buscar o link de download...")
        res, msg_data_full = imap.fetch(latest_msg_id, "(RFC822)")
        raw_email = msg_data_full[0][1]
        msg = email.message_from_bytes(raw_email)
        
        html_body = ""
        for part in msg.walk():
            if part.get_content_type() == "text/html":
                payload = part.get_payload(decode=True)
                html_body = payload.decode(part.get_content_charset() or 'utf-8', errors='replace')
                break
                
        if not html_body:
            print("Corpo HTML não encontrado no e-mail. Pulando extração de link.")
            imap.logout()
            return
            
        # Extrair links usando regex
        hrefs = re.findall(r'href="([^"]+)"', html_body, re.IGNORECASE)
        
        # Filtrar links do SendGrid
        sg_urls = []
        for href in hrefs:
            clean_href = href.strip().replace("&amp;", "&")
            if "sendgrid.net/ls/click" in clean_href and clean_href not in sg_urls:
                sg_urls.append(clean_href)
                
        print(f"Encontrados {len(sg_urls)} links do SendGrid. Identificando o link do relatório...")
        
        download_success = False
        for url in sg_urls:
            try:
                # Fazer requisição GET com stream=True para ler apenas o cabeçalho e seguir redirecionamentos
                res_link = requests.get(url, stream=True, allow_redirects=True, timeout=10)
                final_url = res_link.url
                content_type = res_link.headers.get("Content-Type", "")
                
                # Checar se o link final corresponde a uma planilha excel ou csv
                is_excel = "spreadsheet" in content_type.lower() or "excel" in content_type.lower() or "csv" in content_type.lower()
                is_file_ext = ".xlsx" in final_url.lower() or ".csv" in final_url.lower() or ".xls" in final_url.lower()
                
                if (is_excel or is_file_ext) and not final_url.lower().endswith(".pdf"):
                    print(f"Link de download identificado: {final_url[:120]}...")
                    print("Baixando o arquivo do relatório...")
                    
                    # Fazer o download completo
                    file_res = requests.get(url, allow_redirects=True, timeout=20)
                    dest_path = os.path.join(BASE_DIR, "relatorio_ton.csv")
                    with open(dest_path, "wb") as f:
                        f.write(file_res.content)
                        
                    print(f"Relatório Ton baixado e salvo com sucesso em: {dest_path}")
                    download_success = True
                    break
            except Exception as e:
                print(f"Erro ao verificar/baixar link {url[:50]}: {e}")
                
        if download_success:
            # Salvar o ID do e-mail como processado
            with open(last_processed_path, "w", encoding="utf-8") as f:
                f.write(msg_id)
            
            # Mover o e-mail para a Lixeira (Trash) do Gmail e exclui-lo da caixa de entrada
            try:
                print(f"Movendo e-mail (ID {latest_msg_id.decode()}) para a Lixeira...")
                imap.copy(latest_msg_id, "[Gmail]/Trash")
                imap.store(latest_msg_id, "+FLAGS", "\\Deleted")
                imap.expunge()
                print("E-mail movido para a Lixeira e excluído da Caixa de Entrada.")
            except Exception as e_del:
                print(f"Aviso: Não foi possível excluir o e-mail: {e_del}")
        else:
            print("Nenhum link de download de relatório encontrado nos links do e-mail.")
            
        imap.logout()
    except Exception as e:
        print(f"Erro ao baixar e-mail da Ton: {e}")
        traceback.print_exc()

def run_automation():
    # 0. Tentar baixar o relatório da Ton por e-mail, se disponível
    try:
        baixar_relatorio_ton()
    except Exception as e:
        print(f"Erro não crítico ao baixar e-mail da Ton: {e}")
        traceback.print_exc()

    print("Iniciando automação do relatório PCLAB...")
    with sync_playwright() as p:
        print("Inicializando navegador Chromium...")
        browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-gpu"])
        context = browser.new_context(
            accept_downloads=True,
            locale="pt-BR",
            timezone_id="America/Sao_Paulo"
        )
        page = context.new_page()
        
        # 1. Acessar site e fazer Login
        print("Acessando a página de login...")
        page.goto("https://dashboard.pclab.com.br/Home/Login")
        
        print("Preenchendo formulário de login...")
        page.fill("#chave", "SAOPAULO")
        page.fill("#usuario", "Pedro")
        page.fill("#pww_form", "Pedro1504")
        
        print("Clicando no botão de login...")
        page.click("#js-login-btn")
        page.wait_for_load_state("networkidle")
        
        # 2. Navegar diretamente para o relatório de fluxo de caixa
        print("Navegando para a página do relatório...")
        report_url = "https://dashboard.pclab.com.br/Relatorio?id=205&fkMenu=6&descricao=Relat%C3%B3rio%20-%20Fluxo%20de%20Caixa&sigla=RELFLUXOCAIXA&mostrar=True&arquivo=ReportFluxoCaixa"
        page.goto(report_url)
        
        # 3. Esperar carregar o Stimulsoft
        print("Aguardando carregamento e renderização do Stimulsoft...")
        try:
            page.wait_for_selector("input.stiJsViewerTextBox", timeout=30000)
            time.sleep(2)  # Small buffer to ensure rendering is complete
        except Exception as e:
            print(f"Aviso/Erro ao aguardar inputs do Stimulsoft: {e}")
        
        # 4. Ajustar datas
        inputs = page.query_selector_all("input")
        visible_textboxes = []
        for inp in inputs:
            cls = inp.get_attribute("class") or ""
            if "stiJsViewerTextBox" in cls and inp.is_visible() and inp.is_enabled():
                visible_textboxes.append(inp)
                
        print(f"Encontrados {len(visible_textboxes)} campos de texto visíveis e habilitados.")
        if len(visible_textboxes) < 2:
            raise Exception("Não foi possível encontrar os campos de data no relatório.")
            
        start_date_input = visible_textboxes[0]
        end_date_input = visible_textboxes[1]
        
        today = datetime.now()
        thirty_days_ago = today - timedelta(days=30)
        thirty_days_ago_str = thirty_days_ago.strftime("%d.%m.%Y")
        
        print(f"Preenchendo Data Inicial (30 dias atrás): {thirty_days_ago_str}")
        start_date_input.click()
        start_date_input.press("Control+A")
        start_date_input.press("Backspace")
        start_date_input.fill(thirty_days_ago_str)
        start_date_input.press("Tab")
        
        # 5. Clicar em Enviar
        enviar_btn = None
        for el in page.query_selector_all("div, span, td, a"):
            try:
                if el.is_visible() and el.inner_text().strip() == "Enviar":
                    enviar_btn = el
                    break
            except Exception:
                pass
                
        if not enviar_btn:
            raise Exception("Não foi possível encontrar o botão 'Enviar'.")
            
        print("Clicando no botão 'Enviar'...")
        enviar_btn.click()
        
        # Aguardar dados atualizarem
        print("Aguardando carregamento dos dados atualizados (8 segundos)...")
        time.sleep(8)
        
        # 6. Clicar em Salvar
        salvar_btn = None
        for el in page.query_selector_all("div, span, td, a"):
            try:
                if el.is_visible() and el.inner_text().strip() == "Salvar":
                    salvar_btn = el
                    break
            except Exception:
                pass
                
        if not salvar_btn:
            raise Exception("Não foi possível encontrar o botão 'Salvar'.")
            
        print("Clicando no botão 'Salvar'...")
        salvar_btn.click()
        time.sleep(2)
        
        # 7. Escolher "Arquivo Microsoft Excel..."
        excel_option = None
        for el in page.query_selector_all("div, span, td, a, tr"):
            try:
                if el.is_visible() and el.inner_text().strip() == "Arquivo Microsoft Excel...":
                    excel_option = el
                    break
            except Exception:
                pass
                
        if not excel_option:
            raise Exception("Não foi possível encontrar a opção 'Arquivo Microsoft Excel...' no menu.")
            
        print("Selecionando a exportação para 'Arquivo Microsoft Excel...'...")
        excel_option.click()
        time.sleep(3)
        
        # 8. Clicar em OK no diálogo e capturar o download
        ok_btn = None
        ok_options = []
        for el in page.query_selector_all("div, span, td, button"):
            try:
                if el.is_visible() and el.inner_text().strip() == "OK":
                    ok_options.append(el)
            except Exception:
                pass
                
        if not ok_options:
            raise Exception("Não foi possível encontrar o botão 'OK' de confirmação da exportação.")
            
        ok_btn = ok_options[0]
        
        print("Clicando em 'OK' e iniciando o download...")
        with page.expect_download(timeout=120000) as download_info:
            ok_btn.click()
            
        download = download_info.value
        dest_path = os.path.join(BASE_DIR, "fluxo_caixa.xls")
        
        # Salvar o arquivo
        print(f"Download iniciado. Salvando em: {dest_path}")
        download.save_as(dest_path)
        
        browser.close()

    # 9. Executar upload_receipts.py para subir os dados para o Supabase
    print("Importando e executando upload_receipts.py para subir dados para o Supabase...")
    upload_script = os.path.join(BASE_DIR, "upload_receipts.py")
    if not os.path.exists(upload_script):
        print(f"Aviso: Script de upload não encontrado em {upload_script}. Ignorando etapa de upload.")
        return
        
    try:
        if BASE_DIR not in sys.path:
            sys.path.insert(0, BASE_DIR)
        import upload_receipts
        import importlib
        importlib.reload(upload_receipts)
        upload_receipts.main()
        print("Upload para o Supabase concluído com sucesso!")
    except BaseException as e:
        print(f"Erro ao executar upload_receipts.py: {e}")
        traceback.print_exc()
        raise Exception("Falha ao subir registros para o Supabase através do upload_receipts.py.")

def main_once():
    print("=== Iniciando execução única da automação ===")
    run_automation()
    print("=== Execução única concluída com sucesso! ===")

def main_daemon():
    print("=== Iniciando serviço daemon de atualização do fluxo de caixa ===")
    while True:
        try:
            print(f"Executando atualização às {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}...")
            run_automation()
            print("Atualização concluída com sucesso. Próxima execução em 30 minutos.")
            time.sleep(1800)  # Espera 30 minutos (1800 segundos)
        except Exception as e:
            print(f"Erro na automação: {e}")
            traceback.print_exc()
            print("Tentando novamente em 5 minutos...")
            time.sleep(300)  # Em caso de erro, tenta novamente em 5 minutos (300 segundos)

if __name__ == "__main__":
    if "--once" in sys.argv:
        try:
            main_once()
        except Exception as e:
            print(f"Erro na execução única: {e}")
            traceback.print_exc()
            sys.exit(1)
    else:
        main_daemon()
