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

EXCEL_PATH = "imovdisp.csv"

config = configparser.ConfigParser()
files_read = config.read("config.ini")
if not files_read:
    raise FileNotFoundError("config.ini file not found or could not be read.")

try:
    IMAP_HOST = config["EMAIL"]["IMAP_HOST"]
    SMTP_HOST = config["EMAIL"]["SMTP_HOST"]
    SMTP_PORT = int(config["EMAIL"].get("SMTP_PORT", 587))
except KeyError as e:
    raise KeyError(f"Missing required config key: {e}")

# === FUNÇÕES ===
import os

import csv

import tempfile

def preprocess_csv(input_path):
    temp_file = tempfile.NamedTemporaryFile(delete=False, mode='w', encoding='latin1', newline='')
    with open(input_path, 'r', encoding='latin1', errors='ignore') as infile, temp_file:
        for line in infile:
            # Basic cleanup: remove problematic characters or fix quotes if needed
            cleaned_line = line.replace('\x00', '')  # Remove null bytes if any
            # Additional cleaning rules can be added here
            temp_file.write(cleaned_line)
    return temp_file.name

def carregar_planilha(caminho):
    ext = os.path.splitext(caminho)[1].lower()
    if ext == '.csv':
        cleaned_path = preprocess_csv(caminho)
        try:
            try:
                df = pd.read_csv(cleaned_path, dtype=str, encoding='latin1', delimiter=';', quoting=csv.QUOTE_MINIMAL, on_bad_lines='warn')
                print("DEBUG: CSV loaded with semicolon delimiter and QUOTE_MINIMAL after preprocessing")
            except Exception as e1:
                print(f"DEBUG: Failed to load CSV with semicolon delimiter after preprocessing: {e1}")
                try:
                    df = pd.read_csv(cleaned_path, dtype=str, encoding='latin1', delimiter='\t', quoting=csv.QUOTE_MINIMAL, on_bad_lines='warn')
                    print("DEBUG: CSV loaded with tab delimiter and QUOTE_MINIMAL after preprocessing")
                except Exception as e2:
                    print(f"DEBUG: Failed to load CSV with tab delimiter after preprocessing: {e2}")
                    raise e2
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(cleaned_path, dtype=str, encoding='latin1', delimiter=';', quoting=csv.QUOTE_MINIMAL, on_bad_lines='warn')
                print("DEBUG: CSV loaded with semicolon delimiter, latin1 encoding and QUOTE_MINIMAL after preprocessing")
            except Exception as e3:
                print(f"DEBUG: Failed to load CSV with semicolon delimiter and latin1 encoding after preprocessing: {e3}")
                try:
                    df = pd.read_csv(cleaned_path, dtype=str, encoding='latin1', delimiter='\t', quoting=csv.QUOTE_MINIMAL, on_bad_lines='warn')
                    print("DEBUG: CSV loaded with tab delimiter, latin1 encoding and QUOTE_MINIMAL after preprocessing")
                except Exception as e4:
                    print(f"DEBUG: Failed to load CSV with tab delimiter and latin1 encoding after preprocessing: {e4}")
                    raise e4
    elif ext in ['.xls', '.xlsx']:
        df = pd.read_excel(caminho, dtype=str)
    else:
        raise ValueError(f"Unsupported file extension: {ext}")

    first_col = df.columns[0]
    df[first_col] = df[first_col].str.strip()
    df.set_index(first_col, inplace=True)
    df.index = df.index.astype(str)
    # Convert index to list, strip spaces and quotes from each element, then set back as index
    cleaned_index = [x.strip().strip('\'"') for x in df.index.tolist()]
    df.index = pd.Index(cleaned_index)
    print(f"DEBUG: DataFrame index keys after cleaning: {list(df.index)}")
    print(f"DEBUG: DataFrame index values (first 10) after cleaning: {list(df.index[:10])}")
    return df

def montar_email(codigo, dados):
    return f"""E-MAIL AUTOMÁTICO - por favor não responder! Para mais informações entre em contato pelo número (85) 99984-3733.

Bom dia!

Recebemos um e-mail da ZAP+ informando que você teria interesse em um imóvel que está para locação. Segue abaixo informações do imóvel:

{dados['TIPO DO IMOVEL'].upper()} – Código {codigo}

Endereço: {dados['ENDEREÇO']}.
Aluguel: {dados['ALUGUEL']}
Proprietário: {dados['PROPRIET.']}
Situação: {dados.get('SITUAÇÃO', 'Não informado')}
Inscrição IPTU: {dados.get('INSC_IPTU', 'Não informado')}

E-mail do Proprietário: {dados.get('E-MAIL PROP.', 'Não informado')}

Atenciosamente,
Edilson & Edilia Administração de Imóveis Ltda
https://edilsoneediliaimoveis.com.br/
85 99984-3733
85 3221-6272
"""

import socket

def enviar_email(destinatario, assunto, corpo, smtp_user, smtp_pass):
    msg = MIMEText(corpo)
    msg["Subject"] = assunto
    msg["From"] = smtp_user
    msg["To"] = destinatario

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(smtp_user, smtp_pass)
            server.send_message(msg)
            print(f"E-mail enviado para {destinatario} com sucesso.")
    except (socket.gaierror, ConnectionRefusedError) as e:
        print(f"Erro de conexão SMTP: {e}")
        raise
    except smtplib.SMTPException as e:
        print(f"Erro SMTP: {e}")
        raise

def processar_emails(email_usuario, senha_email, log_callback=None):
    df = carregar_planilha(EXCEL_PATH)
    count_emails = 0

    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    try:
        print(f"Connecting to IMAP server: {IMAP_HOST}")
        mail = imaplib.IMAP4_SSL(IMAP_HOST)
        print(f"Logging in as: {email_usuario}")
        mail.login(email_usuario, senha_email)
        print("Login successful")
        mail.select("inbox")

        result, data = mail.search(None, '(UNSEEN FROM "mateuscad98@gmail.com")')

        for num in data[0].split():
            result, message_data = mail.fetch(num, '(RFC822)')
            raw_email = message_data[0][1]
            mensagem = email.message_from_bytes(raw_email)

            subject = mensagem.get('Subject', '(Sem Assunto)')
            log(f"Lendo e-mail: {subject}")

            try:
                if mensagem.is_multipart():
                    for part in mensagem.walk():
                        if part.get_content_type() == "text/plain":
                            corpo = part.get_payload(decode=True).decode()
                            break
                    else:
                        corpo = ""
                else:
                    corpo = mensagem.get_payload(decode=True).decode()
            except Exception as e:
                log(f"Erro ao decodificar corpo do e-mail: {e}")
                corpo = str(raw_email)

            log(f"Corpo do e-mail recebido:\n{corpo}")

            match_codigo = re.search(r'C[ÓO]D[.:]?\s*0*([0-9A-Za-z-]+)', corpo, re.IGNORECASE)
            match_email = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', corpo)

            if match_codigo and match_email:
                codigo = match_codigo.group(1).strip().lstrip('0')
                destinatario = match_email.group(0).strip()
                log(f"Código interno encontrado: {codigo}")

                if codigo in df.index:
                    dados = df.loc[codigo]
                    if dados.get("DISPONIBILIDADE", "").lower() == "disponível":
                        texto = montar_email(codigo, dados)
                        enviar_email(destinatario, f"Informações do imóvel {codigo}", texto, email_usuario, senha_email)
                        log(f"E-mail enviado para {destinatario} com o assunto: Informações do imóvel {codigo}")
                        mail.store(num, '+FLAGS', '\\Seen')
                        count_emails += 1
                    else:
                        log(f"Imóvel {codigo} indisponível.")
                else:
                    log(f"Código {codigo} não encontrado.")
            else:
                log("Código ou e-mail do cliente não encontrado no corpo da mensagem.")

        mail.logout()
        log(f"Total de e-mails processados com códigos internos: {count_emails}")
    except (socket.gaierror, ConnectionRefusedError) as e:
        log(f"Erro de conexão IMAP: {e}")
        raise
    except imaplib.IMAP4.error as e:
        log(f"Erro IMAP: {e}")
        raise

# === INTERFACE GRÁFICA - QUINTA VERSÃO (08.01.25) ===

def iniciar_interface():
    janela = tk.Tk()
    janela.title("LeadBot - E-mail Automático")
    janela.iconbitmap("leadbot.ico")
    janela.geometry("400x250")
    janela.resizable(True, True)
    fonte = ("Arial", 12)
    janela.config(bg="#333333")

    estado_logado = {"logado": False}
    mail_connection = {"mail": None}

    # Exibição da senha
    def toggle_password():
        if entrada_senha.cget('show') == '':
            entrada_senha.config(show='*')
            btn_toggle.config(image=closed_eye_img)
        else:
            entrada_senha.config(show='')
            btn_toggle.config(image=open_eye_img)

    # Login servidor IMAP
    def tentar_login():
        email_usuario = entrada_email.get()
        senha = entrada_senha.get()

        if not email_usuario or not senha:
            messagebox.showerror("Erro", "Por favor, preencha todos os campos.")
            return

        try:
            email_usuario_stripped = email_usuario.strip()
            senha_stripped = senha.strip()
            print(f"DEBUG: Attempting login with email: '{email_usuario_stripped}' and password length: {len(senha_stripped)}")
            mail = imaplib.IMAP4_SSL(IMAP_HOST)
            mail.login(email_usuario_stripped, senha_stripped)
            mail.logout()
            estado_logado["logado"] = True

            # Remover campos de login e senha
            entrada_email.pack_forget()
            entrada_senha.pack_forget()
            btn_toggle.pack_forget()
            label_email.pack_forget()
            label_senha.pack_forget()
            senha_frame.pack_forget()
            btn_logar.pack_forget()

            # Mostrar mensagem de boas-vindas e área de logs
            label_bem_vindo.config(text=f"Bem vindo, {email_usuario}!")
            label_bem_vindo.pack(pady=(20, 5))
            janela.geometry("600x500")
            text_logs.pack(pady=(10, 5), fill=tk.BOTH, expand=True)
            btn_enviar.pack(side=tk.BOTTOM, fill=tk.X, pady=5)
            btn_enviar.lift()

            janela.bind('<Return>', lambda event: enviar_emails())

        except imaplib.IMAP4.error as e:
            messagebox.showerror("Erro", f"Falha no login: {str(e)}")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")

    def enviar_emails():
        email_usuario = entrada_email.get()
        senha = entrada_senha.get()
        try:

            text_logs.config(state=tk.NORMAL)
            text_logs.delete(1.0, tk.END)

            def log_callback(msg):
                text_logs.config(state=tk.NORMAL)
                text_logs.insert(tk.END, msg + "\n")
                text_logs.see(tk.END)
                text_logs.config(state=tk.DISABLED)

            count_emails = 0
            def processar_emails_com_contagem(email_usuario, senha_email, log_callback=None):
                nonlocal count_emails
                df = carregar_planilha(EXCEL_PATH)
                count_emails = 0

                def log(msg):
                    if log_callback:
                        log_callback(msg)
                    else:
                        print(msg)

                try:
                    mail = imaplib.IMAP4_SSL(IMAP_HOST)
                    mail.login(email_usuario, senha_email)
                    mail.select("inbox")

                    result, data = mail.search(None, '(UNSEEN FROM "mateuscad98@gmail.com")')

                    for num in data[0].split():
                        result, message_data = mail.fetch(num, '(RFC822)')
                        raw_email = message_data[0][1]
                        mensagem = email.message_from_bytes(raw_email)

                        subject = mensagem.get('Subject', '(Sem Assunto)')
                        log(f"Lendo e-mail: {subject}")

                        if mensagem.is_multipart():
                            for part in mensagem.walk():
                                if part.get_content_type() == "text/plain":
                                    corpo = part.get_payload(decode=True).decode()
                                    break
                        else:
                            corpo = mensagem.get_payload(decode=True).decode()

                        log(f"Corpo do e-mail recebido:\n{corpo}")

                        match_codigo = re.search(r'C[ÓO]D[.:]?\s*0*([0-9A-Za-z-]+)', corpo, re.IGNORECASE)
                        match_email = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', corpo)

                        if match_codigo:
                            codigo = match_codigo.group(1).strip()
                            if match_email:
                                destinatario = match_email.group(0).strip()
                                log(f"Código interno encontrado: {codigo}")
                                log(f"E-mail do cliente encontrado: {destinatario}")
                            else:
                                destinatario = None
                                log("E-mail do cliente não encontrado no corpo da mensagem.")

                        codigo_norm = codigo.strip().strip('\'"').upper()
                        df_index_norm = df.index.str.strip().str.strip('\'"').str.upper()
                        log(f"DEBUG: Normalized extracted code: {repr(codigo_norm)}")
                        log(f"DEBUG: Normalized DataFrame index keys: {[repr(i) for i in df_index_norm]}")
                        if codigo_norm in df_index_norm.values:
                            dados = df.loc[df_index_norm == codigo_norm].iloc[0]
                            if dados.get("DISPONIBILIDADE", "").lower() == "disponível":
                                texto = montar_email(codigo, dados)
                                enviar_email(destinatario, f"Informações do imóvel {codigo}", texto)
                                log(f"E-mail enviado para {destinatario} com o assunto: Informações do imóvel {codigo}")
                                mail.store(num, '+FLAGS', '\\Seen')
                                count_emails += 1
                            else:
                                log(f"Imóvel {codigo} indisponível.")
                        else:
                            log(f"Código {codigo} não encontrado.")
                    else:
                        log("Código ou e-mail do cliente não encontrado no corpo da mensagem.")

                    mail.logout()
                    log(f"Total de e-mails processados com códigos internos: {count_emails}")
                except (socket.gaierror, ConnectionRefusedError) as e:
                    log(f"Erro de conexão IMAP: {e}")
                    raise
                except imaplib.IMAP4.error as e:
                    log(f"Erro IMAP: {e}")
                    raise
                return count_emails

            count = processar_emails_com_contagem(email_usuario, senha, log_callback)

            text_logs.config(state=tk.NORMAL)
            text_logs.insert(tk.END, f"Leads processados com sucesso! Total de e-mails enviados: {count}\n")
            text_logs.see(tk.END)
            text_logs.config(state=tk.DISABLED)
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")

    label_email = tk.Label(janela, text="Seu E-mail:", font=fonte, bg="#333333", fg="white")
    label_email.pack(pady=(20, 5))
    entrada_email = tk.Entry(janela, width=40, font=fonte)
    entrada_email.pack()
    
    label_senha = tk.Label(janela, text="Sua Senha:", font=fonte, bg="#333333", fg="white")
    label_senha.pack(pady=(10, 5))
    senha_frame = tk.Frame(janela, bg="#333333")
    senha_frame.pack(pady=(10, 5))

    label_bem_vindo = tk.Label(janela, text="", font=("Arial", 14), bg="#333333", fg="white")

    text_logs = tk.Text(janela, bg="#222222", fg="white", font=("Consolas", 10), state=tk.DISABLED)

    entrada_senha = tk.Entry(senha_frame, width=35, font=fonte, show="*")
    entrada_senha.pack(side=tk.LEFT)

    open_eye_img = tk.PhotoImage(file="open_eye.png")
    closed_eye_img = tk.PhotoImage(file="closed_eye.png")

    btn_toggle = tk.Button(senha_frame, image=closed_eye_img, command=toggle_password, bg="#555555", fg="white", relief=tk.FLAT)
    btn_toggle.pack(side=tk.LEFT, padx=(5, 0))


    btn_logar = tk.Button(janela, text="Logar", font=fonte, command=tentar_login, bg="#4CAF50", fg="white")
    btn_logar.pack(pady=20)

    btn_enviar = tk.Button(janela, text="Enviar E-mails Automáticos", font=fonte, command=enviar_emails, bg="#2196F3", fg="white")

    janela.bind('<Return>', lambda event: tentar_login())

    janela.mainloop()

# === EXECUÇÃO ===
if __name__ == "__main__":
    iniciar_interface()