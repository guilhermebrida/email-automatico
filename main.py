import smtplib
from imaplib import IMAP4_SSL
from poplib import POP3_SSL
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart, MIMEBase
import openpyxl
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from threading import Thread
from email import encoders 
import os
import datetime
import customtkinter

customtkinter.set_appearance_mode("Dark")

sender_email = ''
sender_password = ''
smtp_server = 'smtp.office365.com'
smtp_port = 587


excel_file_path = ''


def addLog(data):
    if os.path.exists('./envio-emails-logs.txt'):
        with open('./envio-emails-logs.txt','a') as log:
            log.write(f"-> {datetime.datetime.now()} - {str(data)}\n")
    else:
        with open('./envio-emails-logs.txt','w') as log:
            log.write(f"-> {datetime.datetime.now()} - {str(data)}\n")

def obter_credenciais():
    try:
        global sender_email, sender_password

        sender_email = entry_email.get()
        sender_password = entry_senha.get()
        print(sender_email, sender_password)

    except Exception as e:
        addLog(f'obter_credenciais - {str(e)}')

def obter_caminho_planilha():
    global excel_file_path, sheet
    excel_file_path = filedialog.askopenfilename(title="Selecionar Planilha")
    workbook = openpyxl.load_workbook(excel_file_path)
    sheet = workbook.active
    

def enviar_emails():
    try:
        dados_emails = [row for row in sheet.iter_rows(min_row=2, values_only=True) if any(row)]
        print(dados_emails)
        for idx, (nome, email,cc, anexo_path, nf) in enumerate(dados_emails):
            print(email)
            body = f'''Olá {nome}, \n
Identificamos pendência da NF {nf}. Segue anexa NF.\n
Informamos que o CL está bloqueado para compras junto a BRS até regularização.\n
Gentileza enviar o MDE com o lançamento para desbloqueio do CL.\n
Aguardamos breve retorno.\n
Qualquer dúvida estamos à disposição.\n
Abraço\n
Atenciosamente, 

        '''

            message = MIMEMultipart()
            message['From'] = sender_email
            message['To'] = email
            # message['To'] = ', '.join([email, cc])
            message['Subject'] = f'''{nome} BLOQUEADO PARA COMPRAS - NF {nf} EM ABERTO'''
            message['Cc'] = cc
            message.attach(MIMEText(body, 'plain'))

            if anexo_path is not None:
                with open(anexo_path, 'rb') as anexo:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(anexo.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename="{anexo_path}"')
                    message.attach(part)

            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.sendmail(sender_email, email, message.as_string())
                addLog(f'enviar_emails - E-mail enviado para {email} - {nome} - {nf}')

            progresso_var.set((idx + 1) / len(dados_emails))
            janela_principal.update()

        addLog('enviar_emails - Envio de e-mails concluído!')
        messagebox.showinfo("Concluído", "Envio de e-mails concluído!")
        
    except Exception as e:
        addLog(f'enviar_emails - {str(e)}')
        messagebox.showerror("Erro", "Erro ao enviar e-mails!")

def iniciar_envio():
    try:
        obter_credenciais()
        global thread_envio
        thread_envio = Thread(target=enviar_emails)
        thread_envio.start()
    except Exception as e:
        addLog(f'iniciar_envio - {str(e)}')
        messagebox.showerror("Erro", "Erro ao iniciar envio de e-mails!") 
    



janela_principal = customtkinter.CTk()
janela_principal.title("Envio de E-mails")

customtkinter.CTkLabel(janela_principal, text="E-mail:").pack()
entry_email = customtkinter.CTkEntry(janela_principal, width=250)
entry_email.pack()

customtkinter.CTkLabel(janela_principal, text="Senha:").pack()
entry_senha = customtkinter.CTkEntry(janela_principal, show="*", width=250)
entry_senha.pack()

btn_selecionar_planilha = customtkinter.CTkButton(janela_principal, text="Selecionar Planilha", command=obter_caminho_planilha)
btn_selecionar_planilha.pack(pady=10)

progresso_var = customtkinter.DoubleVar()
progresso_barra = customtkinter.CTkProgressBar(janela_principal, variable=progresso_var,width=300, mode='determinate')
progresso_barra.pack(pady=10)

btn_iniciar = customtkinter.CTkButton(janela_principal, text="Iniciar Envio", command=iniciar_envio)
btn_iniciar.pack(pady=10)

janela_principal.mainloop()