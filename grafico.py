import pandas as pd
import unicodedata
import pyautogui
import pyperclip
import webbrowser
import win32com.client as win32
import time
import requests

# === PARTE 1: LEITURA E FILTRAGEM ===
print("[1/7] Lendo e filtrando planilha original...")
fonte = r"C:\Users\raiss\OneDrive\Documentos\projetosdados\precos_anp.xlsx"
df = pd.read_excel(fonte, skiprows=9)
df.columns = df.columns.str.upper().str.strip()

def remover_acentos(txt):
    return unicodedata.normalize('NFKD', txt).encode('ASCII', 'ignore').decode('ASCII')

df.columns = [remover_acentos(col) for col in df.columns]
df['PRODUTO'] = df['PRODUTO'].str.upper().str.strip()

if 'DATA DA COLETA' in df.columns:
    df['DATA DA COLETA'] = pd.to_datetime(df['DATA DA COLETA'], errors='coerce')

df_filtrado = df[
    (df['ESTADO'] == 'BAHIA') &
    (df['MUNICIPIO'].str.upper().isin(['SALVADOR', 'LAURO DE FREITAS'])) &
    (df['PRODUTO'].isin(['GASOLINA COMUM', 'ETANOL']))
]

df_final = df_filtrado[['RAZAO', 'ENDERECO', 'MUNICIPIO', 'BAIRRO', 'BANDEIRA', 'PRODUTO', 'PRECO DE REVENDA', 'DATA DA COLETA']].copy()
df_final.rename(columns={'RAZAO': 'POSTO'}, inplace=True)
df_final['DATA DA COLETA'] = pd.to_datetime(df_final['DATA DA COLETA'], errors='coerce')
df_final['DATA DA COLETA'] = df_final['DATA DA COLETA'].dt.strftime('%d/%m/%Y')

print("[2/7] Salvando nova planilha sem senha...")
arquivo_excel_filtrado = r"C:\Users\raiss\OneDrive\Documentos\projetosdados\planilha_filtrada.xlsx"
df_final.to_excel(arquivo_excel_filtrado, index=False)

# === PARTE 2: MELHORES PREÇOS + LOCALIZAÇÃO ===
gasolina = df_final[df_final['PRODUTO'] == 'GASOLINA COMUM']
etanol = df_final[df_final['PRODUTO'] == 'ETANOL']

linha_gasolina = gasolina.loc[gasolina['PRECO DE REVENDA'].idxmin()]
posto_gasolina = linha_gasolina['POSTO']
endereco_gasolina = linha_gasolina['ENDERECO']
bairro_gasolina = linha_gasolina['BAIRRO']
preco_gasolina = linha_gasolina['PRECO DE REVENDA']

linha_etanol = etanol.loc[etanol['PRECO DE REVENDA'].idxmin()]
posto_etanol = linha_etanol['POSTO']
endereco_etanol = linha_etanol['ENDERECO']
bairro_etanol = linha_etanol['BAIRRO']
preco_etanol = linha_etanol['PRECO DE REVENDA']

# Coordenadas via Nominatim
def obter_coordenadas(endereco):
    url = "https://nominatim.openstreetmap.org/search"
    params = {"q": endereco, "format": "json", "limit": 1}
    headers = {"User-Agent": "script-gasolina"}
    try:
        resp = requests.get(url, params=params, headers=headers)
        data = resp.json()
        if data:
            return data[0]["lat"], data[0]["lon"]
    except:
        pass
    return None, None

# === GERA LINK DO WAZE ===
lat_gas, lon_gas = obter_coordenadas(endereco_gasolina)
lat_eta, lon_eta = obter_coordenadas(endereco_etanol)

link_waze_gas = f"https://waze.com/ul?ll={lat_gas},{lon_gas}&navigate=yes" if lat_gas and lon_gas else "Localização não encontrada"
link_waze_eta = f"https://waze.com/ul?ll={lat_eta},{lon_eta}&navigate=yes" if lat_eta and lon_eta else "Localização não encontrada"

# === MENSAGEM FINAL ===
mensagem = (
    f"⛽️ *Resumo da semana:*\n\n"
    f"🚘 Gasolina mais barata:\n"
    f"*{posto_gasolina}* – R$ {preco_gasolina:.2f}\n"
    f"Bairro: {bairro_gasolina}\n"
    f"[Ir no Waze]({link_waze_gas})\n\n"
    f"🍃 Etanol mais barato:\n"
    f"*{posto_etanol}* – R$ {preco_etanol:.2f}\n"
    f"Bairro: {bairro_etanol}\n"
    f"[Ir no Waze]({link_waze_eta})\n\n"
)

# === PARTE 3: COPIA IMAGEM DO EXCEL ===
print("[3/7] Abrindo Excel e copiando intervalo...")
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True
wb = excel.Workbooks.Open(arquivo_excel_filtrado)
ws = wb.Sheets(1)

print("[3.1] Formatando colunas...")
ws.Columns("A:H").AutoFit()
ws.Columns("G").NumberFormat = "#,##0.00"
ws.Columns("H").NumberFormat = "dd/mm/yyyy"

copiado = False
for tentativa in range(5):
    try:
        ws.Range("A1:H29").CopyPicture(Appearance=1, Format=2)
        copiado = True
        break
    except:
        time.sleep(1)

wb.Close(False)
excel.Quit()

# === PARTE 4: ENVIO PELO WHATSAPP WEB ===
link_grupo = "https://chat.whatsapp.com/Hpp4ic4B85pBEau9s5f08F"
print("[4/7] Abrindo grupo no WhatsApp Web...")
webbrowser.open(link_grupo)
time.sleep(10)

print("[5/7] Entrando na conversa...")
pyautogui.click(x=2284, y=130)
time.sleep(10)

print("[6/7] Clicando no grupo...")
pyautogui.click(x=275, y=320)
time.sleep(2)

print("[7/7] Colando imagem...")
pyautogui.hotkey("ctrl", "v")
time.sleep(2)

print("[8/8] Colando mensagem e enviando...")
pyperclip.copy(mensagem)
pyautogui.hotkey("ctrl", "v")
time.sleep(1)
pyautogui.press("enter")
print("✅ Tudo enviado com sucesso!")
