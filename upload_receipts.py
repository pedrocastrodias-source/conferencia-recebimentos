import os
import csv
import xlrd
import openpyxl
import requests
from datetime import datetime
from dotenv import load_dotenv

# Carrega as variáveis de ambiente do arquivo .env
script_dir = os.path.dirname(os.path.abspath(__file__))
env_path = os.path.join(script_dir, '.env')
load_dotenv(env_path)

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

if not SUPABASE_URL or not SUPABASE_KEY:
    print("Erro: SUPABASE_URL ou SUPABASE_KEY não configurados no arquivo .env.")
    exit(1)

# Caminhos dos arquivos
excel_path = os.path.join(script_dir, 'fluxo_caixa.xls')
csv_path = os.path.join(script_dir, 'relatorio_ton.csv')

def parse_date(date_str):
    if not date_str:
        return None
    try:
        # '25/05/2026 10:19' -> 'YYYY-MM-DD HH:MM:SS'
        dt = datetime.strptime(str(date_str).strip(), '%d/%m/%Y %H:%M')
        return dt.strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return None

def parse_money(val_str):
    if not val_str:
        return 0.0
    try:
        # Limpa 'R$', pontos de milhar e substitui a vírgula decimal por ponto
        cleaned = str(val_str).replace('R$', '').replace('.', '').replace(',', '.').strip()
        return float(cleaned)
    except Exception:
        return 0.0

def import_pclab(headers):
    import re
    print("==================================================")
    print("INICIANDO IMPORTAÇÃO: RECEBIMENTOS PCLAB (EXCEL)")
    print("==================================================")
    
    if not os.path.exists(excel_path):
        print(f"Aviso: Arquivo do Excel não encontrado em: {excel_path}. Pulando etapa PCLAB.")
        return

    print(f"Lendo o arquivo Excel: {excel_path}")
    try:
        workbook = xlrd.open_workbook(excel_path)
        sheet = workbook.sheet_by_index(0)
    except Exception as e:
        print(f"Erro ao abrir o arquivo Excel: {e}")
        return

    print(f"Total de linhas no arquivo Excel: {sheet.nrows}")

    records_all = []
    # As linhas de dados começam após o cabeçalho na linha 15 (índice 16)
    for r in range(16, sheet.nrows):
        row_values = sheet.row_values(r)
        if not row_values or len(row_values) <= 26:
            continue
            
        # Parse do ID (Coluna 0)
        raw_id = row_values[0]
        if isinstance(raw_id, float):
            rec_id = int(raw_id)
        elif isinstance(raw_id, int):
            rec_id = raw_id
        else:
            val_str = str(raw_id).strip()
            # Se for vazio ou não contiver apenas dígitos, não é uma linha de dados (pode ser totais ou rodapés)
            if not val_str or not val_str.isdigit():
                continue
            rec_id = int(val_str)
            
        descricao = str(row_values[3]).strip()
        tipo = str(row_values[17]).strip()
        forma_pgto = str(row_values[20]).strip()
        dt_lancamento = parse_date(row_values[22])
        
        try:
            vr_lanc = float(row_values[26])
        except (ValueError, TypeError):
            vr_lanc = 0.0

        records_all.append({
            "id": rec_id,
            "descricao": descricao,
            "tipo": tipo,
            "forma_pgto": forma_pgto,
            "dt_lancamento": dt_lancamento,
            "vr_lanc": vr_lanc
        })

    # Filtragem de estornos (saídas) e entradas correspondentes
    saidas = [r for r in records_all if r["tipo"] == "Saída"]
    entradas = [r for r in records_all if r["tipo"] == "Entrada"]
    
    excluded_entrada_ids = set()
    excluded_saida_ids = set()
    
    def extract_atendimento(desc):
        match = re.search(r'(?:Atendimento|Atend\.?):\s*(\d+)', desc, re.IGNORECASE)
        if match:
            return match.group(1)
        return None

    # Mapeamento rápido de entradas por Atendimento
    entrada_by_atend = {}
    for ent in entradas:
        atend = extract_atendimento(ent["descricao"])
        if atend:
            entrada_by_atend.setdefault(atend, []).append(ent)

    for s in saidas:
        excluded_saida_ids.add(s["id"])
        atend = extract_atendimento(s["descricao"])
        if not atend:
            print(f"Aviso: Não foi possível extrair Atendimento da saída ID {s['id']}.")
            continue
            
        # Procura a entrada correspondente com o mesmo valor e atendimento
        candidates = entrada_by_atend.get(atend, [])
        matched = False
        for cand in candidates:
            if cand["id"] not in excluded_entrada_ids and cand["vr_lanc"] == s["vr_lanc"]:
                excluded_entrada_ids.add(cand["id"])
                print(f"Estorno identificado: Excluindo Entrada ID {cand['id']} correspondente à Saída ID {s['id']} (Atendimento: {atend}, Valor: {s['vr_lanc']})")
                matched = True
                break
        
        if not matched:
            print(f"Aviso: Estorno ID {s['id']} não encontrou uma entrada correspondente de valor {s['vr_lanc']} para atendimento {atend}.")

    # Registros finais válidos após filtragem
    records = [r for r in records_all if r["id"] not in excluded_saida_ids and r["id"] not in excluded_entrada_ids]
    
    print(f"Total de registros originais no Excel: {len(records_all)}")
    print(f"Saídas (estornos) identificadas e excluídas: {len(saidas)}")
    print(f"Entradas correspondentes excluídas: {len(excluded_entrada_ids)}")
    print(f"Total de registros a processar após filtragem: {len(records)}")

    if not records:
        print("Nenhum registro válido encontrado no Excel pós-filtragem.")
        return

    url_get = f"{SUPABASE_URL}/rest/v1/Conferencia_PCLAB?select=id"
    print("Consultando registros já existentes no Supabase (PCLAB)...")
    try:
        response = requests.get(url_get, headers=headers)
        if response.status_code != 200:
            print(f"Erro ao buscar registros no Supabase: {response.status_code} - {response.text}")
            return
        
        existing_ids = {item["id"] for item in response.json()}
        print(f"Registros encontrados no banco de dados (PCLAB): {len(existing_ids)}")
    except Exception as e:
        print(f"Erro na conexão com o Supabase: {e}")
        return

    # Filtragem de duplicados
    new_records = [rec for rec in records if rec["id"] not in existing_ids]
    print(f"Novos registros PCLAB a serem inseridos: {len(new_records)}")
    print(f"Registros PCLAB duplicados ignorados: {len(records) - len(new_records)}")

    if not new_records:
        print("Todos os registros PCLAB já estão cadastrados. Nenhum upload necessário.")
        return

    # Upload em lote
    chunk_size = 500
    url_post = f"{SUPABASE_URL}/rest/v1/Conferencia_PCLAB"
    
    headers_post = headers.copy()
    headers_post["Content-Type"] = "application/json"
    headers_post["Prefer"] = "return=minimal"
    
    uploaded_count = 0
    for i in range(0, len(new_records), chunk_size):
        chunk = new_records[i:i + chunk_size]
        print(f"Enviando lote de {len(chunk)} registros PCLAB...")
        try:
            res = requests.post(url_post, json=chunk, headers=headers_post)
            if res.status_code in [200, 201]:
                uploaded_count += len(chunk)
                print(f"Lote PCLAB enviado com sucesso. Progresso: {uploaded_count}/{len(new_records)}")
            else:
                print(f"Erro ao enviar lote PCLAB (status {res.status_code}): {res.text}")
                return
        except Exception as e:
            print(f"Erro durante o envio: {e}")
            return

    print(f"Importação PCLAB concluída! {uploaded_count} registros inseridos com sucesso.")

def import_maquininha(headers):
    print("\n==================================================")
    print("INICIANDO IMPORTAÇÃO: RELATÓRIO TON MAQUININHA (CSV/XLSX)")
    print("==================================================")

    if not os.path.exists(csv_path):
        print(f"Aviso: Arquivo de relatório Ton não encontrado em: {csv_path}. Pulando etapa Maquininha.")
        return

    print(f"Lendo o arquivo: {csv_path}")
    records = []
    
    # Verifica se o arquivo é um XLSX mascarado de CSV (inicia com bytes ZIP 'PK\x03\x04')
    is_xlsx = False
    try:
        with open(csv_path, "rb") as f:
            sig = f.read(4)
            if sig == b'PK\x03\x04':
                is_xlsx = True
    except Exception as e:
        print(f"Erro ao verificar formato do arquivo: {e}")
        return

    if is_xlsx:
        print("Detectado formato Excel (XLSX) no arquivo. Processando via openpyxl...")
        try:
            with open(csv_path, "rb") as f:
                wb = openpyxl.load_workbook(f, read_only=True)
                sheet = wb.active
                rows = list(sheet.iter_rows(values_only=True))
                
                if len(rows) > 1:
                    header = rows[0]
                    print(f"Cabeçalho Excel identificado: {header}")
                    
                    for row_idx, row in enumerate(rows[1:], start=1):
                        if not row or len(row) < 10:
                            continue
                        
                        rec_id = str(row[0]).strip() if row[0] is not None else ""
                        if not rec_id:
                            continue
                        
                        raw_date = row[1]
                        if isinstance(raw_date, datetime):
                            dt_transacao = raw_date.strftime('%Y-%m-%d %H:%M:%S')
                        else:
                            dt_transacao = parse_date(raw_date)
                            
                        raw_val = row[2]
                        if isinstance(raw_val, (int, float)):
                            vr_recebido = float(raw_val)
                        else:
                            vr_recebido = parse_money(raw_val)
                            
                        metodo_pgto = str(row[9]).strip() if row[9] is not None else ""
                        
                        records.append({
                            "id": rec_id,
                            "dt_transacao": dt_transacao,
                            "vr_recebido": vr_recebido,
                            "metodo_pgto": metodo_pgto
                        })
                wb.close()
        except Exception as e:
            print(f"Erro ao abrir ou processar o arquivo Excel: {e}")
            return
    else:
        print("Detectado formato CSV (texto delimitado). Processando via leitor CSV...")
        try:
            with open(csv_path, mode='r', encoding='latin-1') as f:
                reader = csv.reader(f, delimiter=';')
                header = next(reader)
                print(f"Cabeçalho CSV identificado: {header}")
                
                for row_idx, row in enumerate(reader):
                    if not row or len(row) < 10:
                        continue
                    
                    rec_id = str(row[0]).strip()
                    if not rec_id:
                        continue
                    
                    dt_transacao = parse_date(row[1])
                    vr_recebido = parse_money(row[2])
                    metodo_pgto = str(row[9]).strip()
                    
                    records.append({
                        "id": rec_id,
                        "dt_transacao": dt_transacao,
                        "vr_recebido": vr_recebido,
                        "metodo_pgto": metodo_pgto
                    })
        except Exception as e:
            print(f"Erro ao abrir ou processar o arquivo CSV: {e}")
            return

    print(f"Total de registros identificados no relatório da Maquininha: {len(records)}")
    if not records:
        print("Nenhum registro válido encontrado no relatório da Maquininha.")
        return

    url_get = f"{SUPABASE_URL}/rest/v1/Conferencia_Maquininha?select=id"
    print("Consultando registros já existentes no Supabase (Maquininha)...")
    try:
        response = requests.get(url_get, headers=headers)
        if response.status_code != 200:
            print(f"Erro ao buscar registros no Supabase: {response.status_code} - {response.text}")
            return
        
        existing_ids = {str(item["id"]).strip() for item in response.json()}
        print(f"Registros encontrados no banco de dados (Maquininha): {len(existing_ids)}")
    except Exception as e:
        print(f"Erro na conexão com o Supabase: {e}")
        return

    # Filtragem de duplicados
    new_records = [rec for rec in records if rec["id"] not in existing_ids]
    print(f"Novos registros Maquininha a serem inseridos: {len(new_records)}")
    print(f"Registros Maquininha duplicados ignorados: {len(records) - len(new_records)}")

    if not new_records:
        print("Todos os registros da Maquininha já estão cadastrados. Nenhum upload necessário.")
        return

    # Upload em lote
    chunk_size = 500
    url_post = f"{SUPABASE_URL}/rest/v1/Conferencia_Maquininha"
    
    headers_post = headers.copy()
    headers_post["Content-Type"] = "application/json"
    headers_post["Prefer"] = "return=minimal"
    
    uploaded_count = 0
    for i in range(0, len(new_records), chunk_size):
        chunk = new_records[i:i + chunk_size]
        print(f"Enviando lote de {len(chunk)} registros Maquininha...")
        try:
            res = requests.post(url_post, json=chunk, headers=headers_post)
            if res.status_code in [200, 201]:
                uploaded_count += len(chunk)
                print(f"Lote Maquininha enviado com sucesso. Progresso: {uploaded_count}/{len(new_records)}")
            else:
                print(f"Erro ao enviar lote Maquininha (status {res.status_code}): {res.text}")
                return
        except Exception as e:
            print(f"Erro durante o envio: {e}")
            return

    print(f"Importação Maquininha concluída! {uploaded_count} registros inseridos com sucesso.")

def main():
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}"
    }
    
    # Executa ambas as importações
    import_pclab(headers)
    import_maquininha(headers)

if __name__ == "__main__":
    main()
