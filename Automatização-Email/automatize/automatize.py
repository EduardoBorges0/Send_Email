import pandas as pd
import locale
import smtplib
import email.message
from tkinter import *
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

def on_enter_press(event):
      if event.keysym == "Return":
        enviar_email_gmail()

def calc():
    global html_table, faturamento, formatacao_em_real, key_product, key
    vendas = pd.read_excel(r"Vendas.xlsx")
    faturamento = vendas.groupby('Produto')["Valor Final"].max()
    dic = dict(faturamento)
    series_data = pd.Series(faturamento)
    html_table = series_data.to_frame().to_html(formatters={"Valor Final": "R${:,.2f}".format})
    max_value = max(dic.values())
    formatacao_em_real = locale.currency(float(max_value), grouping=True)
    for key, value in dic.items():
        if value == max_value:
            key_product = key


def enviar_email_gmail():
    calc()
    text_entry = input_entry.get()
    msg = email.message.Message()
    msg['Subject'] = "Relatório Mensal Da Empresa"
    msg['From'] = ""#Your_Email
    msg['To'] = text_entry
    password = ""##Your password, you can to find this password in your account Gmail
    msg.add_header("Content-Type", "text/html")
    corpo = msg.HTMLBody = f"""
    <p>O produto com maior faturamento é {key_product} com o faturamento de {formatacao_em_real}</p>
    {html_table}
    """
    msg.set_payload(corpo)
    try:
        smt = smtplib.SMTP('smtp.gmail.com: 587')
        smt.starttls()
        smt.login(msg['From'], password)
        smt.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    except:
      label.config(text="Erro")
      label.place(x=250, y=130)
    else:
      label.config(text="Enviado")
      label.place(x=250, y=130)

def send_email():
    global input_entry, label
    root = Tk()
    root.title("Envio de Emails")
    root.geometry("600x200")
    Label(root, text= "Digite para qual Email você quer enviar:").place(x=80, y=50)
    input_entry = Entry(root)
    input_entry.place(x=80, y=80, width=350, height=30)
    btn = Button(root, text="Enviar", command=enviar_email_gmail)
    btn.place(x=450, y=80, width=70, height=30)
    root.bind("<KeyPress>", on_enter_press)
    label = Label(root, text="")
    label.pack()
    root.mainloop()
    on_enter_press()


send_email()