# 📊 Preços de Combustíveis ANP – Bahia

Este projeto automatiza a análise e envio semanal dos preços de combustíveis (Gasolina e Etanol) registrados pela ANP para os municípios de Salvador e Lauro de Freitas, na Bahia. O script realiza a leitura da planilha oficial, filtra os dados relevantes, formata a tabela, identifica os melhores preços e envia uma mensagem com imagem diretamente para um grupo no WhatsApp Web.

---

## ⚙️ Funcionalidades

- 📥 Leitura da planilha da ANP (`.xlsx`)
- 🔍 Filtragem por município, estado e tipo de combustível
- 💰 Identificação do posto com o menor preço por tipo de combustível
- 📅 Tratamento automático de datas e preços
- 📸 Geração automática de imagem da tabela (via Excel)
- 💬 Envio automático da imagem + mensagem no WhatsApp Web com `pyautogui`

---

## 🧠 Tecnologias utilizadas

- `pandas`  
- `openpyxl`  
- `pyautogui`  
- `pyperclip`  
- `win32com.client`  
- `webbrowser`  
- `requests`  
- `Pillow (PIL)`  
- `unicodedata`  
- `subprocess`  
- `time`

---

## 🚀 Como executar

1. Instale as dependências:  
   ```bash
   pip install -r requirements.txt
   ```

2. Atualize o caminho da planilha no script `bot_combustiveis.py`, se necessário.

3. Execute o script:  
   ```bash
   python bot_combustiveis.py
   ```

⚠️ **Importante:** o WhatsApp Web deve estar logado no navegador padrão, e o Excel deve estar instalado na máquina (Windows).

---

## 📂 Estrutura do projeto

```bash
📁 precos-combustiveis-anp/
├── bot_combustiveis.py
├── precos_anp.xlsx
├── planilha_filtrada.xlsx
├── requirements.txt
└── README.md
```

---

## 👩‍💻 Autora

**Raissa Mariana Luna**  
🔗 [LinkedIn](https://www.linkedin.com/in/raissa-luna-a0292b1a0/)  
📧 raissalunana@gmail.com  
🌐 [GitHub](https://github.com/Raissaluna3026)

---

Este projeto demonstra:

- Habilidade em manipulação de dados com `pandas`
- Automação de tarefas com `pyautogui` e `win32com`
- Integração entre dados públicos e canais de comunicação (WhatsApp)
- Organização e empacotamento de scripts Python para uso prático

🔧 Melhorias Futuras:

1. Automatizar o download da planilha da ANP
Atualmente, o arquivo .xlsx precisa ser baixado manualmente toda semana. Pretende-se implementar um sistema de web scraping ou verificação automática da URL do arquivo mais recente no portal da ANP.

2. Eliminar dependência do PyAutoGUI
O envio da imagem e da mensagem pelo WhatsApp Web depende de interações com a interface via pyautogui, exigindo que a janela esteja visível e ativa. O ideal seria migrar para integrações via API (ex: WhatsApp Business API) ou usar bibliotecas como selenium ou whatsapp-web.js, que permitem automações mais estáveis e headless.

3. Expansão para todos os bairros de Salvador e Lauro de Freitas
A ANP nem sempre fornece dados de todos os bairros. Pretende-se desenvolver um modelo baseado em geolocalização e interpolação para estimar os postos mais próximos, mesmo que o bairro não esteja listado na base de dados.

