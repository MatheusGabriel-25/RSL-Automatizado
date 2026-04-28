import sys
import subprocess
import os
import re
import time
import urllib.parse

# ==============================================================================
# 1. VERIFICAÇÃO E INSTALAÇÃO DE BIBLIOTECAS
# ==============================================================================
def verificar_e_instalar_bibliotecas():
    print("=" * 60)
    print("🛠️  VERIFICANDO BIBLIOTECAS NECESSÁRIAS...")
    print("=" * 60)
    
    # Adicionamos o 'pyzotero' na lista de dependências
    bibliotecas_necessarias = ['requests', 'pandas', 'openpyxl', 'tqdm', 'pyzotero']
    todas_instaladas = True

    for lib in bibliotecas_necessarias:
        try:
            __import__(lib)
        except ImportError:
            todas_instaladas = False
            print(f"⚠️  Biblioteca '{lib}' não encontrada. Instalando...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", lib])
                print(f"✅ '{lib}' instalada com sucesso!")
            except Exception as e:
                print(f"❌ Erro ao instalar '{lib}': {e}")
                sys.exit(1)

    if not todas_instaladas:
        print("\n🚀 Todas as bibliotecas foram instaladas. Reiniciando o script...")
        os.execv(sys.executable, ['python'] + sys.argv)

verificar_e_instalar_bibliotecas()

import pandas as pd
import requests
from tqdm import tqdm
from pyzotero import zotero

# ==============================================================================
# 2. CONFIGURAÇÕES E FUNÇÕES AUXILIARES
# ==============================================================================
PASTA_DESTINO = "Revisao Sistema da Literatura Artigos em PDF"

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36',
    'Accept': 'application/pdf, application/xhtml+xml, text/html'
}

def limpar_nome_arquivo(titulo):
    titulo = str(titulo)
    titulo_limpo = re.sub(r'[\\/*?:"<>|]', "", titulo)
    return titulo_limpo[:150].strip()

def baixar_pdf(url, caminho_arquivo):
    """Tenta baixar o PDF a partir de uma URL"""
    try:
        response = requests.get(url, headers=HEADERS, timeout=15, stream=True, allow_redirects=True)
        if response.status_code == 200:
            content_type = response.headers.get('Content-Type', '').lower()
            if 'application/pdf' in content_type or url.lower().endswith('.pdf') or 'application/octet-stream' in content_type:
                with open(caminho_arquivo, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)
                return True, "Download direto concluído."
            else:
                return False, f"Protegido/Não é PDF (Content-Type: {content_type})"
        else:
            return False, f"Erro HTTP {response.status_code}"
    except Exception as e:
        return False, f"Erro de conexão: {str(e)}"

# ==============================================================================
# 3. MOTORES DE RESGATE (Open Access)
# ==============================================================================
def buscar_pdf_alternativo(titulo, doi):
    links_encontrados = []
    doi_limpo = str(doi).replace("https://doi.org/", "").strip() if pd.notna(doi) and str(doi).strip() != "" and str(doi).strip() != "N/A" else None
    
    # 1. OpenAlex
    try:
        if doi_limpo:
            url_openalex = f"https://api.openalex.org/works/https://doi.org/{doi_limpo}"
        else:
            titulo_encoded = urllib.parse.quote(titulo)
            url_openalex = f"https://api.openalex.org/works?filter=title.search:{titulo_encoded}"
            
        resp_oa = requests.get(url_openalex, headers=HEADERS, timeout=10).json()
        if doi_limpo and resp_oa.get('open_access', {}).get('is_oa'):
            url_pdf = resp_oa['open_access'].get('oa_url')
            if url_pdf: links_encontrados.append(('OpenAlex', url_pdf))
        elif not doi_limpo and resp_oa.get('results') and len(resp_oa['results']) > 0:
            if resp_oa['results'][0].get('open_access', {}).get('is_oa'):
                url_pdf = resp_oa['results'][0]['open_access'].get('oa_url')
                if url_pdf: links_encontrados.append(('OpenAlex', url_pdf))
    except: pass

    # 2. Semantic Scholar
    try:
        if doi_limpo:
            url_s2 = f"https://api.semanticscholar.org/graph/v1/paper/DOI:{doi_limpo}?fields=title,openAccessPdf"
        else:
            titulo_encoded = urllib.parse.quote(titulo)
            url_s2 = f"https://api.semanticscholar.org/graph/v1/paper/search?query={titulo_encoded}&limit=1&fields=title,openAccessPdf"
            
        resp_s2 = requests.get(url_s2, headers=HEADERS, timeout=10).json()
        if doi_limpo and resp_s2.get('openAccessPdf'):
            url_pdf = resp_s2['openAccessPdf'].get('url')
            if url_pdf: links_encontrados.append(('Semantic Scholar', url_pdf))
        elif not doi_limpo and resp_s2.get('data') and len(resp_s2['data']) > 0:
            if resp_s2['data'][0].get('openAccessPdf'):
                url_pdf = resp_s2['data'][0]['openAccessPdf'].get('url')
                if url_pdf: links_encontrados.append(('Semantic Scholar', url_pdf))
    except: pass

    return links_encontrados

# ==============================================================================
# 4. LÓGICA PRINCIPAL DO PROGRAMA E MENU
# ==============================================================================
def main():
    print("=" * 60)
    print("🤖 BATCH DOWNLOADER V3 - INTEGRAÇÃO ZOTERO & NOTEBOOK LM")
    print("=" * 60)
    
    os.makedirs(PASTA_DESTINO, exist_ok=True)
    print(f"📁 Pasta de destino: '{PASTA_DESTINO}'\n")

    print("Qual é a fonte dos seus artigos?")
    print("[1] Ler da Planilha Excel")
    print("[2] Ler direto da Nuvem do Zotero")
    opcao = input("👉 Escolha uma opção (1 ou 2): ").strip()

    lista_artigos = []

    if opcao == '1':
        arquivo_excel = input("\n📁 Digite o nome da sua planilha Excel (ex: Dataset_Final.xlsx): ").strip()
        try:
            df = pd.read_excel(arquivo_excel)
            colunas = [c.upper() for c in df.columns]
            
            if 'LINK' not in colunas or 'TÍTULO' not in colunas:
                print("❌ As colunas 'LINK' e/ou 'TÍTULO' não foram encontradas na planilha.")
                return
                
            col_link = df.columns[colunas.index('LINK')]
            col_titulo = df.columns[colunas.index('TÍTULO')]
            col_doi = df.columns[colunas.index('DOI')] if 'DOI' in colunas else None
            
            for index, row in df.iterrows():
                if pd.isna(row[col_titulo]): continue
                rank = row['RANK'] if 'RANK' in colunas else index + 1
                lista_artigos.append({
                    'rank': rank,
                    'titulo': row[col_titulo],
                    'link': row[col_link] if pd.notna(row[col_link]) else "",
                    'doi': row[col_doi] if col_doi and pd.notna(row[col_doi]) else ""
                })
        except Exception as e:
            print(f"❌ Erro ao ler a planilha: {e}")
            return

    elif opcao == '2':
        print("\n🔑 Para conectar ao Zotero, você precisa do seu User ID e API Key.")
        print("Obtenha em: https://www.zotero.org/settings/keys")
        library_id = input("ID do Usuário (User ID): ").strip()
        api_key = input("Chave da API (API Key): ").strip()
        
        try:
            print("⏳ Conectando ao Zotero... Isso pode levar alguns segundos.")
            zot = zotero.Zotero(library_id, 'user', api_key)
            # Baixa todos os itens principais (ignora anexos e notas avulsas)
            itens_zotero = zot.everything(zot.top())
            print(f"✅ Sucesso! Encontrados {len(itens_zotero)} registros principais no seu Zotero.")
            
            rank_contador = 1
            for item in itens_zotero:
                data = item.get('data', {})
                tipo = data.get('itemType')
                
                # Foca apenas em documentos científicos
                if tipo not in ['journalArticle', 'conferencePaper', 'thesis', 'report', 'bookSection']:
                    continue
                    
                lista_artigos.append({
                    'rank': rank_contador,
                    'titulo': data.get('title', 'Sem Titulo'),
                    'link': data.get('url', ''),
                    'doi': data.get('DOI', '')
                })
                rank_contador += 1
                
        except Exception as e:
            print(f"❌ Erro de conexão com o Zotero: {e}")
            print("Verifique se o seu ID e Key estão corretos e se você tem internet.")
            return

    else:
        print("❌ Opção inválida. Encerrando.")
        return

    # ==============================================================================
    # PROCESSO DE DOWNLOAD DOS ARTIGOS (PLANILHA OU ZOTERO)
    # ==============================================================================
    sucessos_diretos = 0
    sucessos_resgate = 0
    falhas = []

    print("\n🚀 Iniciando varredura e downloads...\n")
    
    for artigo in tqdm(lista_artigos, desc="Processando Artigos"):
        titulo = artigo['titulo']
        link_original = artigo['link']
        doi = artigo['doi']
        rank = artigo['rank']
        
        nome_arquivo = f"{rank:02d} - {limpar_nome_arquivo(titulo)}.pdf"
        caminho_completo = os.path.join(PASTA_DESTINO, nome_arquivo)
        
        if os.path.exists(caminho_completo):
            sucessos_diretos += 1
            continue
            
        sucesso, mensagem = False, "Link original ausente"
        if link_original and link_original != "":
            sucesso, mensagem = baixar_pdf(link_original, caminho_completo)
        
        if sucesso:
            sucessos_diretos += 1
        else:
            # Tenta o resgate inteligente
            links_alternativos = buscar_pdf_alternativo(titulo, doi)
            resgatado = False
            
            for fonte, link_alt in links_alternativos:
                sucesso_alt, msg_alt = baixar_pdf(link_alt, caminho_completo)
                if sucesso_alt:
                    sucessos_resgate += 1
                    resgatado = True
                    break
                    
            if not resgatado:
                falhas.append({
                    'Rank': rank, 
                    'Título': titulo, 
                    'Link Original': link_original, 
                    'Erro Base': mensagem,
                    'Resgate': 'Não encontrado em bases abertas'
                })
            
        time.sleep(1.5) # Proteção contra bloqueio de IP

    # ==============================================================================
    # RELATÓRIO FINAL
    # ==============================================================================
    print("\n" + "=" * 60)
    print("📊 RELATÓRIO FINAL DE DOWNLOADS")
    print("=" * 60)
    print(f"✅ Baixados pelo link original: {sucessos_diretos}")
    print(f"🛡️ Baixados via RESGATE (API): {sucessos_resgate}")
    print(f"❌ Exigem download manual: {len(falhas)}")
    
    if falhas:
        arquivo_relatorio = "Relatorio_Falhas_Download.txt"
        with open(arquivo_relatorio, 'w', encoding='utf-8') as f:
            f.write("RELATÓRIO DE ARTIGOS PROTEGIDOS (PAYWALL/ANTI-BOT)\n")
            f.write("Aviso: Baixe estes arquivos manualmente logando no site da revista ou universidade.\n")
            f.write("=" * 80 + "\n\n")
            
            for falha in falhas:
                f.write(f"RANK: {falha.get('Rank')}\n")
                f.write(f"TÍTULO: {falha.get('Título')}\n")
                f.write(f"LINK ORIGINAL: {falha.get('Link Original')}\n")
                f.write(f"STATUS RESGATE: {falha.get('Resgate')}\n")
                f.write("-" * 80 + "\n")
                
        print(f"\n⚠️ Um arquivo chamado '{arquivo_relatorio}' foi criado na sua pasta.")

if __name__ == "__main__":
    main()