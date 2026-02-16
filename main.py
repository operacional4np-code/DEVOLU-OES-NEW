import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
from pathlib import Path

# --- CONFIGURAÇÃO DE CAMINHOS ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# NOME ATUALIZADO DA SUA PLANILHA
NOME_PLANILHA = "Dados.xlsx" 
INPUT_EXCEL = os.path.join(BASE_DIR, NOME_PLANILHA)

# Caminho do Modelo dentro da pasta assets
MODELO_PATH = os.path.join(BASE_DIR, "assets", "modelo_protocolo.png")

# Onde os resultados serão salvos (Pasta Downloads do seu PC)
DOWNLOADS_PATH = Path.home() / "Downloads" / "Protocolos_Gerados"

def gerar_protocolos():
    # Cria a pasta de destino se não existir
    if not DOWNLOADS_PATH.exists():
        DOWNLOADS_PATH.mkdir(parents=True, exist_ok=True)

    # Verifica se a planilha 'dados.xlsx' existe na pasta
    if not os.path.exists(INPUT_EXCEL):
        print(f"❌ Erro: Não encontrei o arquivo '{NOME_PLANILHA}' nesta pasta.")
        print(f"Caminho tentado: {INPUT_EXCEL}")
        return

    try:
        print(f"⏳ Lendo dados de {NOME_PLANILHA}...")
        df = pd.read_excel(INPUT_EXCEL)
        
        for index, row in df.iterrows():
            with Image.open(MODELO_PATH).convert("RGB") as img:
                draw = ImageDraw.Draw(img)
                
                # Tenta carregar Arial, se não conseguir usa a fonte padrão
                try:
                    fonte = ImageFont.truetype("arial.ttf", 22)
                except:
                    fonte = ImageFont.load_default()

                # --- PREENCHIMENTO ---
                # Ajuste os nomes entre [' '] para bater exatamente com o topo da sua tabela
                draw.text((800, 48), str(row['protocolo']), fill="black", font=fonte)
                draw.text((100, 145), str(row['cliente']), fill="black", font=fonte)
                draw.text((150, 242), str(row['nota_fiscal']), fill="black", font=fonte)
                draw.text((550, 242), str(row['cte']), fill="black", font=fonte)
                draw.text((100, 310), str(row['data']), fill="black", font=fonte)
                draw.text((100, 450), str(row['nome_recebedor']), fill="black", font=fonte)

                # --- SALVAMENTO ---
                nome_arq = f"Protocolo_{row['protocolo']}.png"
                img.save(DOWNLOADS_PATH / nome_arq)
                print(f"✅ Protocolo de {row['cliente']} gerado!")

        print(f"\n🚀 Prontinho! Vá até sua pasta de Downloads para ver os arquivos.")

    except Exception as e:
        print(f"❌ Ocorreu um erro inesperado: {e}")

if __name__ == "__main__":
    gerar_protocolos()
