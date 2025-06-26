from io import BytesIO
from PIL import Image
import pandas as pd
import unicodedata
import pyautogui
import pyperclip
import webbrowser
import win32com.client as win32
import os
import subprocess
import time

# === PARTE 1: LEITURA E FILTRAGEM ===
print("[1/7] Lendo e filtrando planilha original...")
fonte = r"C:\Users\raiss\OneDrive\Documentos\projetosdados\precos_anp.xlsx"
df = pd.read_excel(fonte, skiprows=9)
df.columns = df.columns.str.upper().str.strip()

def remover_acentos(txt):
    return unicodedata.normalize('NFKD', txt).encode('ASCII', 'ignore').decode('ASCII')

df.columns = [remover_acentos(col) for col in df.columns]

# Força a conversão da coluna de data (se necessário)
if 'DATA DA COLETA' in df.columns:
    df['DATA DA COLETA'] = pd.to_datetime(df['DATA DA COLETA'], errors='coerce')


df_filtrado = df[
    (df['ESTADO'] == 'BAHIA') &
    (df['MUNICIPIO'].str.upper().isin(['SALVADOR', 'LAURO DE FREITAS'])) &
    (df['PRODUTO'].isin(['GASOLINA COMUM', 'ETANOL']))
]

df_final = df_filtrado[['RAZAO', 'MUNICIPIO', 'BAIRRO', 'BANDEIRA', 'PRODUTO', 'PRECO DE REVENDA', 'DATA DA COLETA']].copy()
df_final.rename(columns={'RAZAO': 'POSTO'}, inplace=True)

# 🔧 Corrige a coluna de data explicitamente
df_final['DATA DA COLETA'] = pd.to_datetime(df_final['DATA DA COLETA'], errors='coerce')



print("[2/7] Salvando nova planilha sem senha...")
arquivo_excel_filtrado = r"C:\Users\raiss\OneDrive\Documentos\projetosdados\planilha_filtrada.xlsx"
df_final.to_excel(arquivo_excel_filtrado, index=False)

# === PARTE 2: EXTRAI MELHORES PREÇOS POR COMBUSTÍVEL ===
gasolina = df_final[df_final['PRODUTO'] == 'GASOLINA COMUM']
etanol = df_final[df_final['PRODUTO'] == 'ETANOL']

linha_gasolina = gasolina.loc[gasolina['PRECO DE REVENDA'].idxmin()]
posto_gasolina = linha_gasolina['POSTO']
bairro_gasolina = linha_gasolina['BAIRRO']
preco_gasolina = linha_gasolina['PRECO DE REVENDA']

linha_etanol = etanol.loc[etanol['PRECO DE REVENDA'].idxmin()]
posto_etanol = linha_etanol['POSTO']
bairro_etanol = linha_etanol['BAIRRO']
preco_etanol = linha_etanol['PRECO DE REVENDA']




mensagem = (
    f"⛽️ *Resumo da semana:*\n\n"
    f"Gasolina mais barata: *{posto_gasolina}* - R$ {preco_gasolina:.2f}\n localizado em: {bairro_gasolina}"
    f"Etanol mais barato: *{posto_etanol}* - R$ {preco_etanol:.2f}\n localizado em: {bairro_etanol}\n"
    f"Segue a lista completa na imagem abaixo 👇"
)

# === PARTE 3: ABRE EXCEL, FORMATA E COPIA TABELA COMO IMAGEM ===
print("[3/7] Abrindo Excel e copiando intervalo...")
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True
wb = excel.Workbooks.Open(arquivo_excel_filtrado)
ws = wb.Sheets(1)

print("[3.1] Formatando colunas...")
ws.Columns("A:H").AutoFit()
ws.Columns("G").NumberFormat = "#,##0.00"
ws.Columns("H").NumberFormat = "dd/mm/yyyy"

# Copia a imagem
copiado = False
tentativas = 0
while not copiado and tentativas < 5:
    try:
        ws.Range("A1:H29").CopyPicture(Appearance=1, Format=2)
        copiado = True
    except Exception as e:
        tentativas += 1
        print(f"Tentativa {tentativas} falhou: {e}")
        time.sleep(1)

wb.Close(False)
excel.Quit()

# === PARTE 4: ABRE WHATSAPP WEB E ENVIA ===
link_grupo = "https://chat.whatsapp.com/Hpp4ic4B85pBEau9s5f08F"

print("[4/7] Abrindo o link do grupo...")
webbrowser.open(link_grupo)
time.sleep(10)

print("[5/7] Clicando no botão 'Entrar na conversa'...")
pyautogui.click(x=2284, y=130)  # ajuste se necessário
time.sleep(10)

print("[6/7] Clicando no grupo para garantir o foco...")
pyautogui.click(x=275, y=320)  # ajuste se necessário
time.sleep(2)

print("[7/7] Colando imagem...")
pyautogui.hotkey("ctrl", "v")
time.sleep(2)

print("[8/8] Colando mensagem e enviando...")
pyperclip.copy(mensagem)
pyautogui.hotkey("ctrl", "v")
time.sleep(1)
pyautogui.press("enter")
print("✅ Imagem e mensagem enviadas com sucesso.")
