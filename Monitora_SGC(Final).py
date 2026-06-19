import pandas as pd
import requests
import json
import os
import sys
from datetime import datetime, timedelta
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2.service_account import Credentials

print("Iniciando o Monitora SGC...")

# Descobre a pasta onde o programa (ou o .exe) está rodando
pasta_do_programa = os.path.dirname(os.path.abspath(sys.argv[0]))

# Define os caminhos relativos de forma dinâmica e segura
caminho_token = os.path.join(pasta_do_programa, "meu_token.txt")
caminho_arquivo_mestre = os.path.join(pasta_do_programa, "Espelho_itens_Mestre.xlsx")
caminho_credenciais = os.path.join(pasta_do_programa, "credenciais.json")

# SISTEMA DE LOGIN / PRIMEIRO ACESSO
if os.path.exists(caminho_token):
    # Se o arquivo já existir, o robô apenas lê o token
    with open(caminho_token, "r", encoding="utf-8") as arquivo:
        token = arquivo.read().strip()
else:
    # Se o arquivo não existir, e for a primeira vez que alguem estiver usando
    print("\n" + "="*50)
    print("BEM-VINDO! CONFIGURAÇÃO DE PRIMEIRO ACESSO")
    print("="*50)
    print("Parece que é a sua primeira vez rodando o sistema.")
    token = input("Por favor, cole o seu Token do Big Data e aperte ENTER:\n> ").strip()
    
    # O robô cria o arquivo de texto e guarda a chave lá dentro
    with open(caminho_token, "w", encoding="utf-8") as arquivo:
        arquivo.write(token)
        
    print("\nSucesso! O seu token foi salvo no arquivo 'meu_token.txt'.")
    print("Você não precisará fazer isso novamente nas próximas vezes.\n")


print("Conectando com a API...")

url = "https://www.bigdata.pe.gov.br/api/buscar"

headers = {
    "Authorization": token,
    "offset": "0",
}

Respostas = requests.get(url, headers=headers)
Respostas.raise_for_status()  # Verifica se a requisição foi bem-sucedida
df = pd.DataFrame(Respostas.json())

print("importação concluída, tratando dados JSON...")

# 1. traduzindo os dois json de uma vez
df['lotes_json'] = df['lotes_json'].apply(lambda x: json.loads(x) if isinstance(x, str) else x)
df['pncp_json'] = df['pncp_json'].apply(lambda x: json.loads(x) if isinstance(x, str) else [])

# 2. explodindo a primeira camada (lotes)
df = df.explode('lotes_json')
df['itens_separados'] = df['lotes_json'].apply(lambda d: d.get('itens', []) if isinstance(d, dict) else [])
df = df.explode('itens_separados')
df = df.dropna(subset=['itens_separados'])

print("Transformando chaves do lote em colunas")

# 3. pandas para json (lotes)
df_novas_colunas = df['itens_separados'].apply(pd.Series)
df = pd.concat([df.drop(columns=['lotes_json', 'itens_separados']), df_novas_colunas], axis=1)

print("iniciando o cruzamento de dados com o pncp")

# 4. a nova logica do pncp
def extrair_pncp_correto(lista_pncp, posicao_alvo):
    if isinstance(lista_pncp, list):
        for item in lista_pncp:
            if item.get('pncp_position_item') == posicao_alvo:
                return item
    return {}

df['pncp_item_exato'] = df.apply(lambda row: extrair_pncp_correto(row.get('pncp_json', []), row.get('position_item')), axis=1)
df_pncp_colunas = df['pncp_item_exato'].apply(pd.Series)
df_pncp_colunas = df_pncp_colunas.add_prefix("PNCP - ")
df = pd.concat([df, df_pncp_colunas], axis=1)

print("Formatando a coluna de Histórico...")

# 1. Traduzindo o json do histórico para uma lista de dicionários (ou uma lista vazia, caso esteja vazio ou seja nulo)
df['historico_json'] = df['historico_json'].apply(lambda x: json.loads(x) if isinstance(x, str) else [])

# 2. Criando a função de compilação do histórico, que transforma a lista de eventos em um texto formatado
def compilar_historico(lista_eventos):
    if not isinstance(lista_eventos, list) or len(lista_eventos) == 0:
        return "Sem histórico"
        
    linhas_historico = []
    for evento in lista_eventos:
        if not isinstance(evento, dict):
            continue
            
        data_bruta = evento.get("created_at", "")
        if "T" in data_bruta:
            data_iso = data_bruta.split("T")[0] 
            partes = data_iso.split("-")
            if len(partes) == 3:
                data_formatada = f"{partes[2]}-{partes[1]}-{partes[0]}"
            else:
                data_formatada = data_iso
        else:
            data_formatada = "Data desconhecida"
            
        acao = evento.get("action", "Ação desconhecida")
        descricao = evento.get("description")
        
        if descricao is None or str(descricao).strip() in ["", "null"]:
            descricao_formatada = "(sem descrição)"
        else:
            descricao_formatada = str(descricao).strip()
            
        linha = f"{data_formatada} - {acao} - {descricao_formatada}"
        linhas_historico.append(linha)
        
    return "\n".join(linhas_historico)

df['Histórico do Processo'] = df['historico_json'].apply(compilar_historico)

print("Limpando colunas desnecessarias e temporarias")

colunas_para_dropar = ['item_id', 'position_item', 'lote_id', 'nome_lote', 
                       'pncp_json', 'pncp_item_exato', 'pncp_position_item', 'historico_json']
df = df.drop(columns=colunas_para_dropar, errors='ignore')

df_final = df.copy().reset_index(drop=True)

print("removendo caracteres ilegais para o excel")

for coluna in df_final.select_dtypes(include=['object', 'string']).columns:
    df_final.loc[:, coluna] = df_final[coluna].replace(r'[\x00-\x09\x0B-\x0C\x0E-\x1F\x7F]', '', regex=True)

print("Analisando as datas históricas do PNCP...")

# 1. Como a API pode ter um delay vamos carimbar a data de ontem para os casos que não tiverem data oficial, mas apresentarem movimentação de status
data_ontem = (datetime.now() - timedelta(days=1)).strftime('%d/%m/%Y')

# 2. Criamos uma chave unica para cruzar os dados históricos de Situação e Homologação, baseada no número do processo + posição do item
df_final['Chave_Unica'] = (
    df_final['process_number'].astype(str).str.strip().str.replace('.0', '', regex=False) + "_" + 
    df_final['PNCP - pncp_position_item'].astype(str).str.strip().str.replace('.0', '', regex=False)
)

df_final['Data da Situação'] = ""
df_final['Primeira Data de Homologação'] = ""


if os.path.exists(caminho_arquivo_mestre):
    try:
        df_antigo = pd.read_excel(caminho_arquivo_mestre)
        df_antigo['Chave_Unica'] = (
            df_antigo['process_number'].astype(str).str.strip().str.replace('.0', '', regex=False) + "_" + 
            df_antigo['PNCP - pncp_position_item'].astype(str).str.strip().str.replace('.0', '', regex=False)
        )
        
        if 'Data da Situação' in df_antigo.columns:
            df_antigo_sit = df_antigo.dropna(subset=['Data da Situação'])
            df_antigo_sit = df_antigo_sit[df_antigo_sit['Data da Situação'].astype(str).str.strip() != ""]
            df_antigo_sit = df_antigo_sit.drop_duplicates(subset=['Chave_Unica'])
            mapa_sit = dict(zip(df_antigo_sit['Chave_Unica'], df_antigo_sit['Data da Situação']))
            df_final['Data da Situação'] = df_final['Chave_Unica'].map(mapa_sit)
            df_final['Data da Situação'] = df_final['Data da Situação'].fillna("").astype(str).str.strip()
            
        if 'Primeira Data de Homologação' in df_antigo.columns:
            df_antigo_hom = df_antigo.dropna(subset=['Primeira Data de Homologação'])
            df_antigo_hom = df_antigo_hom[df_antigo_hom['Primeira Data de Homologação'].astype(str).str.strip() != ""]
            df_antigo_hom = df_antigo_hom.drop_duplicates(subset=['Chave_Unica'])
            mapa_hom = dict(zip(df_antigo_hom['Chave_Unica'], df_antigo_hom['Primeira Data de Homologação']))
            df_final['Primeira Data de Homologação'] = df_final['Chave_Unica'].map(mapa_hom)
            df_final['Primeira Data de Homologação'] = df_final['Primeira Data de Homologação'].fillna("").astype(str).str.strip()
            
        print("Memória aplicada! Histórico de Situação e de Homologação preservados.")
    except Exception as e:
        print(f"Aviso ao ler histórico: {e}. Prosseguindo com colunas vazias.")


# 4. aplicamos as atualizações apenas para os casos que não tiverem data oficial, mas apresentarem movimentação de status

nome_coluna_pncp = 'PNCP - data_homologado'

# A. Identifica se a API trouxe uma data oficial válida nesta consulta
if nome_coluna_pncp in df_final.columns:
    condicao_tem_data_api = (df_final[nome_coluna_pncp].notna()) & (df_final[nome_coluna_pncp].astype(str).str.strip() != "")
    datas_formatadas_api = pd.to_datetime(df_final[nome_coluna_pncp], errors='coerce').dt.strftime('%d/%m/%Y')
else:
    condicao_tem_data_api = pd.Series(False, index=df_final.index)
    datas_formatadas_api = pd.Series("", index=df_final.index)


# --- REGRA A: Data da Situação (Histórico de Movimentação de Status) ---

# 1. Verifica se a linha possui qualquer situação (texto) informada pela API
condicao_tem_situacao = (df_final['PNCP - situacao_compra'].notna()) & (df_final['PNCP - situacao_compra'].astype(str).str.strip() != '')

# 2. TRAVA 1: Garante que o status NÃO seja "Em andamento"
condicao_nao_em_andamento = ~df_final['PNCP - situacao_compra'].astype(str).str.contains('Em andamento', case=False, na=False)

# 3. TRAVA 2: Garante que a API Não mandou uma data oficial de homologação
condicao_sem_data_api = ~condicao_tem_data_api

# 4. Verifica se a nossa coluna histórica de Situação está vazia
condicao_sem_data_sit = (df_final['Data da Situação'] == "") | (df_final['Data da Situação'] == "nan")

# Carimba a data_ontem apenas se passar por todas as travas juntas
df_final.loc[condicao_tem_situacao & condicao_nao_em_andamento & condicao_sem_data_api & condicao_sem_data_sit, 'Data da Situação'] = data_ontem


# --- REGRA B: Primeira Data de Homologação (Cofre da Data Oficial do PNCP) ---

# Verifica se a nossa coluna histórica de Homologação está vazia
condicao_historico_homol_vazio = (df_final['Primeira Data de Homologação'] == "") | (df_final['Primeira Data de Homologação'] == "nan")

# Só grava na coluna se o histórico estiver vazio e se a API trouxer a data oficial
df_final.loc[condicao_historico_homol_vazio & condicao_tem_data_api, 'Primeira Data de Homologação'] = datas_formatadas_api


# --- REGRA C: Limpeza Forçada para Itens "Em andamento" ---

# Se o status atual for "Em andamento", limpamos as duas colunas por garantia.

# Isso blinda o painel caso um item mude de status de volta para a fase ativa.
condicao_em_andamento = df_final['PNCP - situacao_compra'].astype(str).str.contains('Em andamento', case=False, na=False)

df_final.loc[condicao_em_andamento, 'Data da Situação'] = ""
df_final.loc[condicao_em_andamento, 'Primeira Data de Homologação'] = ""

print("Formatando colunas de data para o padrão brasileiro...")
colunas_para_formatar = [
    'Data da criação', 'Data da abertura', 'created_at', 'updated_at', 'received_at',
    'stage_started_at', 'opening_published_at', 'proposal_opened_at', 'award_published_at',
    'result_at', 'clearance_at', 'priority_level_at', 'finalized_at', 
    'atual_entrada_neste_estagio', 'PNCP - data_homologado'
] 

for coluna in colunas_para_formatar:
    if coluna in df_final.columns:
        df_final[coluna] = pd.to_datetime(df_final[coluna], errors='coerce').dt.strftime('%d/%m/%Y')
        df_final[coluna] = df_final[coluna].fillna("")

df_final = df_final.drop(columns=['Chave_Unica'])

print("Exportando para Excel...")
df_final.to_excel(caminho_arquivo_mestre, index=False)

print(f"Processo local concluído! O arquivo mestre foi atualizado com {len(df_final)} linhas analisadas.")

# Vamos perguntar ao usuário se ele deseja fazer o upload automático para o Google Sheets, ou se prefere manter apenas a versão local em Excel. A decisão é sua!
resposta = input("\nVocê deseja fazer o upload e atualizar o painel no Google Sheets? (S/N): ").strip().upper()

if resposta == 'S':
    print("Iniciando o upload automático para o Google Sheets...")

    # ID Oculto para o github
    ID_PLANILHA_GOOGLE = "INSIRA_O_ID_DA_SUA_PLANILHA_AQUI"
    SCOPES = ['https://www.googleapis.com/auth/drive']

    try:
        credenciais = Credentials.from_service_account_file(caminho_credenciais, scopes=SCOPES)
        servico = build('drive', 'v3', credentials=credenciais)
        
        media = MediaFileUpload(
            caminho_arquivo_mestre, 
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            resumable=True
        )
        
        print("Enviando dados e convertendo para formato Google Sheets...")
        
        atualizacao = servico.files().update(
            fileId=ID_PLANILHA_GOOGLE,
            media_body=media,
            fields='id'
        ).execute()
        
        print(f"Sucesso Total! Planilha do Google Sheets (ID: {atualizacao.get('id')}) atualizada automaticamente na nuvem.")

    except Exception as e:
        print(f"Erro ao subir para o Google Sheets: {e}\n(Verifique se as credenciais e o ID da planilha estão corretos).")

else:
    print("\nUpload cancelado pelo usuário. Os dados foram salvos apenas localmente na sua máquina.")
    print("Encerrando o programa.")