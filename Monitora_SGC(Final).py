import pandas as pd
import requests
import json
import os
import sys
from datetime import datetime, timedelta
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2.service_account import Credentials

# ==============================================================================
# 1. CONFIGURAÇÃO E AUTENTICAÇÃO
# ==============================================================================
def obter_configuracoes():
    """Define caminhos e gerencia o token de acesso do usuário."""
    pasta_do_programa = os.path.dirname(os.path.abspath(sys.argv[0]))
    caminhos = {
        'token': os.path.join(pasta_do_programa, "meu_token.txt"),
        'mestre': os.path.join(pasta_do_programa, "Espelho_itens_Mestre.xlsx"),
        'credenciais': os.path.join(pasta_do_programa, "credenciais.json")
    }

    if os.path.exists(caminhos['token']):
        with open(caminhos['token'], "r", encoding="utf-8") as arquivo:
            token = arquivo.read().strip()
    else:
        print("\n" + "="*50 + "\nBEM-VINDO! CONFIGURAÇÃO DE PRIMEIRO ACESSO\n" + "="*50)
        token = input("Por favor, cole o seu Token do Big Data e aperte ENTER:\n> ").strip()
        with open(caminhos['token'], "w", encoding="utf-8") as arquivo:
            arquivo.write(token)
        print("\nSucesso! O seu token foi salvo no arquivo 'meu_token.txt'.\n")
        
    return caminhos, token


# ==============================================================================
# 2. EXTRAÇÃO DE DADOS (EXTRACT)
# ==============================================================================
def extrair_dados_api(token):
    """Conecta na API do Estado, baixa os dados e remove duplicatas brutas."""
    print("Conectando com a API...")
    url = "https://www.bigdata.pe.gov.br/api/buscar"
    headers = {"Authorization": token, "offset": "0"}

    Respostas = requests.get(url, headers=headers)
    Respostas.raise_for_status()
    df = pd.DataFrame(Respostas.json())

    print("Importação concluída. Removendo duplicatas...")
    tamanho_original = len(df)
    df = df.drop_duplicates(subset=['id'], keep='first').reset_index(drop=True)
    
    if tamanho_original != len(df):
        print(f"{tamanho_original - len(df)} processos duplicados removidos.")
    else:
        print("Nenhum processo duplicado removido.")
        
    return df


# ==============================================================================
# 3. TRANSFORMAÇÃO DE DADOS (TRANSFORM)
# ==============================================================================
def compilar_historico(lista_eventos):
    """Função auxiliar para formatar o JSON de histórico em texto."""
    if not isinstance(lista_eventos, list) or len(lista_eventos) == 0:
        return "Sem histórico"
        
    linhas_historico = []
    for evento in lista_eventos:
        if not isinstance(evento, dict): continue
            
        data_bruta = evento.get("created_at", "")
        data_formatada = f"{data_bruta.split('T')[0].split('-')[2]}-{data_bruta.split('T')[0].split('-')[1]}-{data_bruta.split('T')[0].split('-')[0]}" if "T" in data_bruta and len(data_bruta.split('T')[0].split('-')) == 3 else data_bruta.split('T')[0] if "T" in data_bruta else "Data desconhecida"
        
        acao = evento.get("action", "Ação desconhecida")
        descricao = evento.get("description")
        descricao_formatada = "(sem descrição)" if descricao is None or str(descricao).strip() in ["", "null"] else str(descricao).strip()
            
        linhas_historico.append(f"{data_formatada} - {acao} - {descricao_formatada}")
    return "\n".join(linhas_historico)


def extrair_pncp_correto(lista_pncp, posicao_alvo):
    """Função auxiliar para puxar o item exato do PNCP."""
    if isinstance(lista_pncp, list):
        for item in lista_pncp:
            if item.get('pncp_position_item') == posicao_alvo:
                return item
    return {}


def transformar_dados(df, caminho_mestre):
    """Desempacota JSONs, aplica regras de negócio e cruza com a base histórica."""
    print("Tratando dados JSON e explodindo lotes...")
    df['lotes_json'] = df['lotes_json'].apply(lambda x: json.loads(x) if isinstance(x, str) else x)
    df['pncp_json'] = df['pncp_json'].apply(lambda x: json.loads(x) if isinstance(x, str) else [])

    df = df.explode('lotes_json')
    df['itens_separados'] = df['lotes_json'].apply(lambda d: d.get('itens', []) if isinstance(d, dict) else [])
    df = df.explode('itens_separados').dropna(subset=['itens_separados'])

    df_novas_colunas = df['itens_separados'].apply(pd.Series)
    df = pd.concat([df.drop(columns=['lotes_json', 'itens_separados']), df_novas_colunas], axis=1)

    print("Cruzando dados com o PNCP e formatando Histórico...")
    df['pncp_item_exato'] = df.apply(lambda row: extrair_pncp_correto(row.get('pncp_json', []), row.get('position_item')), axis=1)
    df_pncp_colunas = df['pncp_item_exato'].apply(pd.Series).add_prefix("PNCP - ")
    df = pd.concat([df, df_pncp_colunas], axis=1)

    df['historico_json'] = df['historico_json'].apply(lambda x: json.loads(x) if isinstance(x, str) else [])
    df['Histórico do Processo'] = df['historico_json'].apply(compilar_historico)

    colunas_para_dropar = ['item_id', 'position_item', 'lote_id', 'nome_lote', 'pncp_json', 'pncp_item_exato', 'pncp_position_item', 'historico_json']
    df_final = df.drop(columns=colunas_para_dropar, errors='ignore').copy().reset_index(drop=True)

    for coluna in df_final.select_dtypes(include=['object', 'string']).columns:
        df_final.loc[:, coluna] = df_final[coluna].replace(r'[\x00-\x09\x0B-\x0C\x0E-\x1F\x7F]', '', regex=True)

    print("Aplicando regras de Datas Históricas (Memória)...")
    df_final['Chave_Unica'] = df_final['process_number'].astype(str).str.strip().str.replace('.0', '', regex=False) + "_" + df_final['PNCP - pncp_position_item'].astype(str).str.strip().str.replace('.0', '', regex=False)
    df_final['Data da Situação'] = ""
    df_final['Primeira Data de Homologação'] = ""

    if os.path.exists(caminho_mestre):
        try:
            df_antigo = pd.read_excel(caminho_mestre)
            df_antigo['Chave_Unica'] = df_antigo['process_number'].astype(str).str.strip().str.replace('.0', '', regex=False) + "_" + df_antigo['PNCP - pncp_position_item'].astype(str).str.replace('.0', '', regex=False)
            
            for col in ['Data da Situação', 'Primeira Data de Homologação']:
                if col in df_antigo.columns:
                    df_temp = df_antigo.dropna(subset=[col])
                    df_temp = df_temp[df_temp[col].astype(str).str.strip() != ""].drop_duplicates(subset=['Chave_Unica'])
                    mapa = dict(zip(df_temp['Chave_Unica'], df_temp[col]))
                    df_final[col] = df_final['Chave_Unica'].map(mapa).fillna("").astype(str).str.strip()
            print("Memória aplicada com sucesso.")
        except Exception as e:
            print(f"Aviso ao ler histórico: {e}")

    # Regras de Negócio A, B e C
    nome_coluna_pncp = 'PNCP - data_homologado'
    condicao_tem_data_api = (df_final[nome_coluna_pncp].notna()) & (df_final[nome_coluna_pncp].astype(str).str.strip() != "") if nome_coluna_pncp in df_final.columns else pd.Series(False, index=df_final.index)
    datas_formatadas_api = pd.to_datetime(df_final[nome_coluna_pncp], errors='coerce').dt.strftime('%d/%m/%Y') if nome_coluna_pncp in df_final.columns else pd.Series("", index=df_final.index)

    condicao_tem_situacao = (df_final['PNCP - situacao_compra'].notna()) & (df_final['PNCP - situacao_compra'].astype(str).str.strip() != '')
    condicao_nao_em_andamento = ~df_final['PNCP - situacao_compra'].astype(str).str.contains('Em andamento', case=False, na=False)
    condicao_em_andamento = ~condicao_nao_em_andamento
    
    df_final.loc[condicao_tem_situacao & condicao_nao_em_andamento & ~condicao_tem_data_api & (df_final['Data da Situação'] == ""), 'Data da Situação'] = (datetime.now() - timedelta(days=1)).strftime('%d/%m/%Y')
    df_final.loc[(df_final['Primeira Data de Homologação'] == "") & condicao_tem_data_api, 'Primeira Data de Homologação'] = datas_formatadas_api
    
    df_final.loc[condicao_em_andamento, ['Data da Situação', 'Primeira Data de Homologação']] = ""

    colunas_para_formatar = ['Data da criação', 'Data da abertura', 'created_at', 'updated_at', 'received_at', 'stage_started_at', 'opening_published_at', 'proposal_opened_at', 'award_published_at', 'result_at', 'clearance_at', 'priority_level_at', 'finalized_at', 'atual_entrada_neste_estagio', 'PNCP - data_homologado'] 
    for coluna in colunas_para_formatar:
        if coluna in df_final.columns:
            df_final[coluna] = pd.to_datetime(df_final[coluna], errors='coerce').dt.strftime('%d/%m/%Y').fillna("")

    return df_final.drop(columns=['Chave_Unica'])


# ==============================================================================
# 4. CARGA DE DADOS (LOAD) E UPLOAD
# ==============================================================================
def salvar_e_upar_planilha(df, caminhos):
    """Salva o Excel localmente e gerencia o envio para o Google Sheets."""
    print("Exportando para Excel local...")
    df.to_excel(caminhos['mestre'], index=False)
    print(f"Processo local concluído! Arquivo atualizado com {len(df)} linhas.")

    resposta = input("\nVocê deseja fazer o upload para o Google Sheets? (S/N): ").strip().upper()
    if resposta == 'S':
        print("Iniciando upload...")
        ID_PLANILHA_GOOGLE = "INSIRA_O_ID_DA_SUA_PLANILHA_AQUI"
        try:
            credenciais = Credentials.from_service_account_file(caminhos['credenciais'], scopes=['https://www.googleapis.com/auth/drive'])
            servico = build('drive', 'v3', credentials=credenciais)
            media = MediaFileUpload(caminhos['mestre'], mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
            
            atualizacao = servico.files().update(fileId=ID_PLANILHA_GOOGLE, media_body=media, fields='id').execute()
            print(f"Sucesso! Planilha (ID: {atualizacao.get('id')}) atualizada na nuvem.")
        except Exception as e:
            print(f"Erro ao subir para o Google Sheets: {e}")
    else:
        print("Upload cancelado. Dados salvos apenas localmente.")


# ==============================================================================
# 5. ORQUESTRADOR (MAIN)
# ==============================================================================
def main():
    """Função principal que dita a ordem de execução do programa."""
    print("Iniciando o Monitora SGC...\n")
    
    # 1. Setup
    caminhos, token = obter_configuracoes()
    
    # 2. Extract
    df_bruto = extrair_dados_api(token)
    
    # 3. Transform
    df_tratado = transformar_dados(df_bruto, caminhos['mestre'])
    
    # 4. Load
    salvar_e_upar_planilha(df_tratado, caminhos)
    
    print("\nPrograma encerrado.")

# Isso garante que o código só rode se você executar este arquivo diretamente
if __name__ == "__main__":
    main()