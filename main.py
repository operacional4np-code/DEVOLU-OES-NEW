import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
from pathlib import Path

# --- CONFIGURAÇÃO DE CAMINHOS DINÂMICOS ---
# Define a pasta onde o script está sendo executado
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Caminho da Planilha (deve estar na mesma pasta que este script)
INPUT_EXCEL = os.path.join(BASE_DIR, "requirements.txt")

# Caminho do Modelo (dentro da pasta assets)
MODELO_PATH = os.path.join(BASE_DIR, "assets", "modelo_protocolo.png")

# Pasta de Destino: Downloads do usuário
DOWNLOADS_PATH = Path.home() / "Downloads" / "Protocolos_Gerados"

def verificar_arquivos():
    """Verifica se todos os ingredientes estão na mesa antes de começar."""
    erro = False
    if not os.path.exists(INPUT_EXCEL):
        print(f"❌ ERRO: Planilha '{INPUT_EXCEL}' não encontrada.")
        erro = True
    if not os.path.exists(MODELO_PATH):
        print(f"❌ ERRO: Imagem modelo '{MODELO_PATH}' não encontrada.")
        erro = True
    
    if erro:
        print("\n💡 DICA: Verifique se os nomes dos arquivos estão idênticos e na pasta correta.")
        return False
    return True

def gerar_protocolos():
    # 1. Cria a pasta nos Downloads se não existir
    if not DOWNLOADS_PATH.exists():
        DOWNLOADS_PATH.mkdir(parents=True, exist_ok=True)

    if not verificar_arquivos():
        return

    try:
        # 2. Carrega a planilha
        print("📊 Lendo planilha...")
        df = pd.read_excel(INPUT_EXCEL)
        
        # 3. Processa cada linha
        for index, row in df.iterrows():
            with Image.open(MODELO_PATH).convert("RGB") as img:
                draw = ImageDraw.Draw(img)
                
                # Tenta carregar a fonte Arial, senão usa a básica do sistema
                try:
                    fonte = ImageFont.truetype("arial.ttf", 22)
                except:
                    fonte = ImageFont.load_default()

                # --- PREENCHIMENTO DOS DADOS (Ajuste X e Y se necessário) ---
                draw.text((800, 48), str(row['protocolo']), fill="black", font=fonte)
                draw.text((100, 145), str(row['cliente']), fill="black", font=fonte)
                draw.text((150, 242), str(row['nota_fiscal']), fill="black", font=fonte)
                draw.text((550, 242), str(row['cte']), fill="black", font=fonte)
                draw.text((100, 310), str(row['data']), fill="black", font=fonte)
                draw.text((100, 450), str(row['nome_recebedor']), fill="black", font=fonte)

                # --- SALVAMENTO ---
                nome_arquivo = f"Protocolo_{row['protocolo']}.png"
                img.save(DOWNLOADS_PATH / nome_arquivo)
                print(f"✅ Gerado: {nome_arquivo}")

        print(f"\n🚀 Sucesso! Todos os arquivos estão em: {DOWNLOADS_PATH}")

    except Exception as e:
        print(f"❌ Erro durante o processamento: {e}")

if __name__ == "__main__":
    gerar_protocolos()
