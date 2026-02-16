import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
from pathlib import Path

# 1. Configuração de Caminhos
DOWNLOADS_PATH = Path.home() / "Downloads" / "Protocolos_Gerados"
INPUT_EXCEL = "CONTROLE_DEVOLUCOES.xlsx"
MODELO_PATH = "assets/modelo_protocolo.png"

# Criar a pasta no seu Downloads
if not DOWNLOADS_PATH.exists():
    DOWNLOADS_PATH.mkdir(parents=True, exist_ok=True)

def gerar_protocolos():
    try:
        # Carregar planilha
        df = pd.read_excel(INPUT_EXCEL)
        
        for index, row in df.iterrows():
            with Image.open(MODELO_PATH).convert("RGB") as img:
                draw = ImageDraw.Draw(img)
                
                # Tentar carregar fonte padrão do Windows, se não, usa a básica
                try:
                    fonte = ImageFont.truetype("arial.ttf", 22)
                except:
                    fonte = ImageFont.load_default()

                # --- PREENCHIMENTO DOS CAMPOS ---
                # Protocolo (Topo Direita)
                draw.text((800, 48), str(row['protocolo']), fill="black", font=fonte)
                # Cliente
                draw.text((100, 145), str(row['cliente']), fill="black", font=fonte)
                # Nota Fiscal
                draw.text((150, 242), str(row['nota_fiscal']), fill="black", font=fonte)
                # CTE
                draw.text((550, 242), str(row['cte']), fill="black", font=fonte)
                # Data
                draw.text((100, 310), str(row['data']), fill="black", font=fonte)
                # Dados do Recebedor (Nome e RG)
                draw.text((100, 450), str(row['nome_recebedor']), fill="black", font=fonte)

                # --- SALVAR ---
                nome_arquivo = f"Protocolo_{row['protocolo']}.png"
                img.save(DOWNLOADS_PATH / nome_arquivo)
                print(f"✅ Protocolo de {row['cliente']} salvo em Downloads!")

    except Exception as e:
        print(f"❌ Ocorreu um erro: {e}")

if __name__ == "__main__":
    gerar_protocolos()
