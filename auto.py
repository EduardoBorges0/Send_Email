import pandas as pd
import smtplib
import email.message
from tkinter import messagebox, ttk
import customtkinter


def on_enter_press(event):
    if event.keysym == "Return":
        enviar_email_gmail()


def calc():
    global html_table, faturamento, max_value, key_product, key, vendas
    input_month_get = month_entry.get().title()
    try:
        vendas = pd.read_excel(f"arquivos_excel/Vendas-{input_month_get}.xlsx")
        faturamento = vendas.groupby('Produto')["Valor Final"].max()
        dic = dict(faturamento)
        series_data = pd.Series(faturamento)
        html_table = series_data.to_frame().to_html(formatters={"Valor Final": "R${:,.2f}".format})
        max_value = max(dic.values())
        for key, value in dic.items():
            if value == max_value:
                key_product = key
    except FileNotFoundError:
        label.configure(text="")
        label.place(x=250, y=10)
        messagebox.showerror("Erro", "Não existe arquivo para esse mês")


def enviar_email_gmail():
    calc()
    emails = pd.read_excel(r"arquivos_excel/Emails_Vendas.xlsx")
    progress = ttk.Progressbar(root, orient="horizontal", length=320, mode="determinate")
    progress.place(x=160, y=280)
    recipient_emails = emails['Emails:'].tolist()
    msg = email.message.Message()
    msg['Subject'] = "Relatório Mensal Da Empresa"
    msg['From'] = "testeempresa496@gmail.com"
    password = "unsq wkay biqw zccd"
    msg.add_header("Content-Type", "text/html")
    corpo = msg.HTMLBody = f"""
    <p>O produto com maior faturamento é {key_product} com o faturamento de {max_value}</p>
    {html_table}
    """
    msg.set_payload(corpo)
    progress['maximum'] = len(recipient_emails)
    try:
        smt = smtplib.SMTP('smtp.gmail.com', port=587)
        smt.starttls()
        smt.login(msg['From'], password)
        for index, recipient_email in enumerate(recipient_emails):
            label.configure(text="")
            label.place(x=230, y=250)
            smt.sendmail(msg['From'], recipient_email, msg.as_string().encode('utf-8'))
            progress['value'] = index + 1
            root.update_idletasks()
    except smtplib.SMTPRecipientsRefused:
        messagebox.showerror("Erro", "Email não encontrado")
    else:
        label.configure(text="Enviado")
        label.place(x=230, y=250)


def send_email():
    global input_entry, label, month_entry, root, label_placeholder, progress
    root = customtkinter.CTk()
    root.geometry("500x300")
    root.title("Envio de Emails")
    month_entry = customtkinter.CTkEntry(root, placeholder_text="Digite o mês da tabela de vendas", width=210,
                                         height=40)
    month_entry.place(x=150, y=80)
    btn = customtkinter.CTkButton(root, text="Enviar", width=210, height=35, command=enviar_email_gmail)
    btn.place(x=150, y=150)
    root.bind("<KeyPress>", on_enter_press)
    label = customtkinter.CTkLabel(root, text="")
    label.pack()
    root.mainloop()


send_email()
