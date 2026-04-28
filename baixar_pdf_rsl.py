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
    
    bibliotecas_necessarias = ['requests', 'pandas', 'openpyxl', 'tqdm']
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
# 3. MOTORES DE RESGATE (APIs de Open Access)
# ==============================================================================
def buscar_pdf_alternativo(titulo, doi):
    """Busca links de PDF em bases de dados abertas usando DOI ou Título"""
    links_encontrados = []
    
    # Limpa o DOI caso exista
    doi_limpo = str(doi).replace("https://doi.org/", "").strip() if pd.notna(doi) and str(doi).strip() != "" and str(doi).strip() != "N/A" else None
    
    # 1. Tentar OpenAlex (Excelente agregador de Acesso Aberto)
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
            melhor_resultado = resp_oa['results'][0]
            if melhor_resultado.get('open_access', {}).get('is_oa'):
                url_pdf = melhor_resultado['open_access'].get('oa_url')
                if url_pdf: links_encontrados.append(('OpenAlex', url_pdf))
    except:
        pass # Ignora erros de API e tenta a próxima

    # 2. Tentar Semantic Scholar
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
    except:
        pass

    return links_encontrados

# ==============================================================================
# 4. LÓGICA PRINCIPAL DO PROGRAMA
# ==============================================================================
def main():
    print("=" * 60)
    print("🤖 BATCH DOWNLOADER V2 - COM RESGATE OPEN ACCESS")
    print("=" * 60)
    
    os.makedirs(PASTA_DESTINO, exist_ok=True)
    print(f"📁 Pasta de destino: '{PASTA_DESTINO}'\n")

    arquivo_excel = input("📁 Digite o nome da sua planilha Excel (ex: Dataset_Final.xlsx): ").strip()
    
    try:
        df = pd.read_excel(arquivo_excel)
    except Exception as e:
        print(f"❌ Erro ao ler a planilha: {e}")
        return

    colunas = [c.upper() for c in df.columns]
    
    if 'LINK' not in colunas or 'TÍTULO' not in colunas:
        print("❌ As colunas 'LINK' e/ou 'TÍTULO' não foram encontradas na planilha.")
        return
    
    coluna_link_real = df.columns[colunas.index('LINK')]
    coluna_titulo_real = df.columns[colunas.index('TÍTULO')]
    coluna_doi_real = df.columns[colunas.index('DOI')] if 'DOI' in colunas else None

    sucessos_diretos = 0
    sucessos_resgate = 0
    falhas = []

    print("\n🚀 Iniciando varredura e downloads...\n")
    
    for index, row in tqdm(df.iterrows(), total=len(df), desc="Processando Artigos"):
        titulo = row[coluna_titulo_real]
        link_original = row[coluna_link_real]
        doi = row[coluna_doi_real] if coluna_doi_real else None
        rank = row['RANK'] if 'RANK' in colunas else index + 1
        
        if pd.isna(titulo):
            continue
            
        nome_arquivo = f"{rank:02d} - {limpar_nome_arquivo(titulo)}.pdf"
        caminho_completo = os.path.join(PASTA_DESTINO, nome_arquivo)
        
        if os.path.exists(caminho_completo):
            sucessos_diretos += 1
            continue
            
        # 1ª Tentativa: Link Original
        sucesso, mensagem = False, "Link original inválido"
        if pd.notna(link_original) and str(link_original).strip() != "":
            sucesso, mensagem = baixar_pdf(link_original, caminho_completo)
        
        if sucesso:
            sucessos_diretos += 1
        else:
            # 2ª Tentativa: Módulo de Resgate (Semantic Scholar / OpenAlex)
            links_alternativos = buscar_pdf_alternativo(titulo, doi)
            resgatado = False
            
            for fonte, link_alt in links_alternativos:
                sucesso_alt, msg_alt = baixar_pdf(link_alt, caminho_completo)
                if sucesso_alt:
                    sucessos_resgate += 1
                    resgatado = True
                    break # Se conseguiu baixar, sai do loop de tentativas
                    
            if not resgatado:
                falhas.append({
                    'Rank': rank, 
                    'Título': titulo, 
                    'Link Original': link_original, 
                    'Erro Base': mensagem,
                    'Resgate': 'Não encontrado em bases abertas'
                })
            
        time.sleep(1.5)

    # ==============================================================================
    # 5. GERAÇÃO DE RELATÓRIO
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