import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
from pathlib import Path

# --- CONFIGURAÇÃO DE CAMINHOS NO CODESPACES ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# NOME EXATO DA PLANILHA (Verifique se é 'dados.xlsx' ou 'Dados.xlsx')
NOME_PLANILHA = "dados.xlsx" 
INPUT_EXCEL = os.path.join(BASE_DIR, NOME_PLANILHA)

# Caminho do Modelo dentro da pasta assets
MODELO_PATH = os.path.join(BASE_DIR, "assets", "modelo_protocolo.png")

# No Codespaces, vamos salvar em uma pasta dentro do próprio projeto para você poder baixar
OUTPUT_DIR = os.path.join(BASE_DIR, "protocolos_prontos")

def gerar_protocolos():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    # Verificação amigável
    if not os.path.exists(INPUT_EXCEL):
        print(f"❌ Erro: Não encontrei '{NOME_PLANILHA}' em {BASE_DIR}")
        print(f"Arquivos que eu encontrei aqui: {os.listdir(BASE_DIR)}")
        return

    try:
        print(f"⏳ Lendo {NOME_PLANILHA}...")
        df = pd.read_excel(INPUT_EXCEL)
        
        for index, row in df.iterrows():
            with Image.open(MODELO_PATH).convert("RGB") as img:
                draw = ImageDraw.Draw(img)
                fonte = ImageFont.load_default()

                # Preenchimento
                draw.text((800, 48), str(row['protocolo']), fill="black", font=fonte)
                draw.text((100, 145), str(row['cliente']), fill="black", font=fonte)
                draw.text((150, 242), str(row['nota_fiscal']), fill="black", font=fonte)
                draw.text((550, 242), str(row['cte']), fill="black", font=fonte)
                draw.text((100, 310), str(row['data']), fill="black", font=fonte)
                draw.text((100, 450), str(row['nome_recebedor']), fill="black", font=fonte)

                nome_arq = f"Protocolo_{row['protocolo']}.png"
                img.save(os.path.join(OUTPUT_DIR, nome_arq))
                print(f"✅ Gerado: {nome_arq}")

        print(f"\n🚀 Sucesso! Veja os arquivos na pasta 'protocolos_prontos' na barra lateral esquerda.")

    except Exception as e:
        print(f"❌ Erro no processamento: {e}")

if __name__ == "__main__":
    gerar_protocolos()
