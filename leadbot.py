#!/usr/bin/env python3
"""
leadbot_imap.py

Script completo para:
1. Ler e-mails da caixa de entrada via IMAP
2. Extrair código do imóvel e e-mail do cliente do corpo da mensagem
3. Buscar informações do imóvel em um Excel
4. Enviar e-mail automático com os dados usando SMTP
5. Marcar o e-mail original como lido
"""

import tkinter as tk
from tkinter import messagebox
import imaplib
import email
import re
import pandas as pd
import smtplib
from email.mime.text import MIMEText

# === CONFIGURAÇÕES ===

import configparser

config = configparser.ConfigParser()
files_read = config.read("config.ini")
if not files_read:
    raise FileNotFoundError("config.ini file not found or could not be read.")

try:
    IMAP_HOST = config["EMAIL"]["IMAP_HOST"]
    IMAP_USER = config["EMAIL"]["IMAP_USER"]
    IMAP_PASS = config["EMAIL"]["IMAP_PASS"]

    SMTP_HOST = config["EMAIL"]["SMTP_HOST"]
    SMTP_USER = config["EMAIL"]["SMTP_USER"]
    SMTP_PASS = config["EMAIL"]["SMTP_PASS"]
    SMTP_PORT = int(config["EMAIL"].get("SMTP_PORT", 587))
except KeyError as e:
    raise KeyError(f"Missing required config key: {e}")

EXCEL_PATH = "imoveis_lista.xlsx"

# === FUNÇÕES ===

def carregar_planilha(caminho):
    df = pd.read_excel(caminho, dtype=str)
    df.set_index('CÓDIGO', inplace=True)
    return df

def montar_email(codigo, dados):
    return f"""E-MAIL AUTOMÁTICO - por favor não responder! Para mais informações entre em contato pelo número (85) 99984-3733.

Bom dia!

Recebemos um e-mail da ZAP+ informando que você teria interesse em um imóvel que está para locação. Segue abaixo informações do imóvel:

{dados['TIPO'].upper()} – Código {codigo}

Endereço: {dados['ENDEREÇO']}.
Aluguel: {dados['ALUGUEL']}
Tipo de imóvel: {dados['TIPO']}

CARACTERÍSTICAS:
{dados['CARACTERÍSTICAS']}

Valor do IPTU: {dados['IPTU']} – referente ao ano de 2025.
Área aproximada: {dados['M²']} m²

Referências: {dados['REFERÊNCIAS']}
Chaves: {dados['CHAVES']}
Garantias de locação: {dados['GARANTIAS']}

Disponibilidade: {dados.get('DISPONIBILIDADE', 'Não informado')}

Atenciosamente,
Edilson & Edilia Administração de Imóveis Ltda
https://edilsoneediliaimoveis.com.br/
85 99984-3733
85 3221-6272
"""

import socket
from oauth2_helper import OAuth2Client, authenticate
import base64

# Add your OAuth2 client configuration here or load from config
OAUTH2_CONFIG = {
    'client_id': 'YOUR_CLIENT_ID',
    'client_secret': 'YOUR_CLIENT_SECRET',
    'auth_uri': 'https://accounts.google.com/o/oauth2/auth',
    'token_uri': 'https://oauth2.googleapis.com/token',
    'redirect_uri': 'http://localhost:8080/',
    'scope': 'https://mail.google.com/'
}

oauth_client = OAuth2Client(**OAUTH2_CONFIG)

def get_oauth2_access_token():
    token = oauth_client.get_access_token()
    if not token:
        token_data = authenticate(oauth_client)
        token = token_data.get('access_token')
    return token

def generate_oauth2_string(username, access_token):
    auth_string = f"user={username}\1auth=Bearer {access_token}\1\1"
    return base64.b64encode(auth_string.encode()).decode()

def enviar_email(destinatario, assunto, corpo):
    access_token = get_oauth2_access_token()
    if not access_token:
        raise Exception("Failed to obtain OAuth2 access token")

    msg = MIMEText(corpo)
    msg["Subject"] = assunto
    msg["From"] = SMTP_USER
    msg["To"] = destinatario

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            auth_string = generate_oauth2_string(SMTP_USER, access_token)
            server.docmd('AUTH', 'XOAUTH2 ' + auth_string)
            server.send_message(msg)
            print(f"E-mail enviado para {destinatario} com sucesso.")
    except (socket.gaierror, ConnectionRefusedError) as e:
        print(f"Erro de conexão SMTP: {e}")
        raise
    except smtplib.SMTPException as e:
        print(f"Erro SMTP: {e}")
        raise

def processar_emails(email_usuario, senha_email):
    df = carregar_planilha(EXCEL_PATH)

    access_token = get_oauth2_access_token()
    if not access_token:
        raise Exception("Failed to obtain OAuth2 access token")

    try:
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        auth_string = generate_oauth2_string(email_usuario, access_token)
        mail.authenticate('XOAUTH2', lambda x: auth_string)
        mail.select("inbox")

        result, data = mail.search(None, '(UNSEEN FROM "noreply@comunica.zapimoveis.com.br")')

        for num in data[0].split():
            result, message_data = mail.fetch(num, '(RFC822)')
            raw_email = message_data[0][1]
            mensagem = email.message_from_bytes(raw_email)

            if mensagem.is_multipart():
                for part in mensagem.walk():
                    if part.get_content_type() == "text/plain":
                        corpo = part.get_payload(decode=True).decode()
                        break
            else:
                corpo = mensagem.get_payload(decode=True).decode()

            # Extrair código
            match_codigo = re.search(r'C[ÓO]D[.:]?\s*([0-9A-Za-z-]+)', corpo, re.IGNORECASE)
            match_email = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', corpo)

            if match_codigo and match_email:
                codigo = match_codigo.group(1).strip()
                destinatario = match_email.group(0).strip()

                if codigo in df.index:
                    dados = df.loc[codigo]
                    if dados.get("DISPONIBILIDADE", "").lower() == "disponível":
                        texto = montar_email(codigo, dados)
                        enviar_email(destinatario, f"Informações do imóvel {codigo}", texto)
                        mail.store(num, '+FLAGS', '\\Seen')
                    else:
                        print(f"Imóvel {codigo} indisponível.")
                else:
                    print(f"Código {codigo} não encontrado.")
            else:
                print("Código ou e-mail do cliente não encontrado no corpo da mensagem.")

        mail.logout()
    except (socket.gaierror, ConnectionRefusedError) as e:
        print(f"Erro de conexão IMAP: {e}")
        raise
    except imaplib.IMAP4.error as e:
        print(f"Erro IMAP: {e}")
        raise

# === INTERFACE GRÁFICA - QUARTA VERSÃO (07.01.25) ===

def iniciar_interface():
    def executar():
        email_usuario = entrada_email.get()
        senha = entrada_senha.get()

        if not email_usuario or not senha:
            messagebox.showerror("Erro", "Por favor, preencha todos os campos.")
            return

        try:
            processar_emails(email_usuario, senha)
            messagebox.showinfo("Sucesso", "Leads processados com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")

    janela = tk.Tk()
    janela.title("LeadBot - E-mail Automático")
    
    # Definindo o ícone (substitua o caminho pelo caminho do seu arquivo .ico)
    janela.iconbitmap("leadbot.ico")

    janela.geometry("400x250")
    janela.resizable(False, False)

    fonte = ("Arial", 12)


    janela.config(bg="#333333")

    # Títulos e Labels com cores mais claras
    tk.Label(janela, text="Seu E-mail:", font=fonte, bg="#333333", fg="white").pack(pady=(20, 5))
    entrada_email = tk.Entry(janela, width=40, font=fonte)
    entrada_email.pack()

    tk.Label(janela, text="Sua Senha:", font=fonte, bg="#333333", fg="white").pack(pady=(10, 5))
    entrada_senha = tk.Entry(janela, width=40, font=fonte, show="*")
    entrada_senha.pack()

    # Botão com fundo escuro e texto claro
    tk.Button(janela, text="Executar", font=fonte, command=executar, bg="#4CAF50", fg="white").pack(pady=20)

    def toggle_password():
        if entrada_senha.cget('show') == '':
            entrada_senha.config(show='*')
            btn_toggle.config(text='Mostrar')
        else:
            entrada_senha.config(show='')
            btn_toggle.config(text='Ocultar')

    btn_toggle = tk.Button(janela, text="Mostrar", font=("Arial", 9), command=toggle_password, bg="#555555", fg="white")
    btn_toggle.pack(pady=(0, 10))

    # Bind Enter key to executar function
    janela.bind('<Return>', lambda event: executar())

    janela.mainloop()

# === EXECUÇÃO ===
if __name__ == "__main__":
    iniciar_interface()
