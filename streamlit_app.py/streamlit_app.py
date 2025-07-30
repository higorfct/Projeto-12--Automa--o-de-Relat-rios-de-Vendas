import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import os
from fpdf import FPDF
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText

# Configuração da página
st.set_page_config(page_title="Relatório de Vendas", layout="centered")
st.title("📊 Automação de Relatório de Vendas")
st.write("Envie uma planilha `.xlsx` para gerar e enviar o relatório por e-mail.")

# Upload do arquivo
uploaded_file = st.file_uploader("Selecione o arquivo de vendas (.xlsx)", type="xlsx")

# Inputs de e-mail
email_remetente = st.text_input("Seu e-mail (Gmail)", "")
senha_app = st.text_input("Senha de app do Gmail", type="password")
email_destino = st.text_input("Destinatário", "")

# Quando o botão for clicado
if uploaded_file and email_remetente and senha_app and email_destino:
    if st.button("📤 Gerar Relatório e Enviar"):
        try:
            # Ler a planilha
            df = pd.read_excel(uploaded_file)

            # Criar pasta de saída
            os.makedirs("output", exist_ok=True)

            # --- RESUMOS ---
            resumo_vendedor = df.groupby("Vendedor")["Valor Total"].sum().reset_index()
            resumo_estado = df.groupby("Estado")["Valor Total"].sum().reset_index()
            resumo_produto = df.groupby("Produto")["Valor Total"].sum().reset_index()

            # --- GRÁFICOS ---
            def salvar_grafico(data, categoria, filename):
                plt.figure(figsize=(6, 4))
                plt.bar(data[categoria], data["Valor Total"], color="mediumseagreen")
                plt.title(f"Total Vendido por {categoria}")
                plt.ylabel("R$")
                plt.xticks(rotation=45)
                plt.tight_layout()
                path = f"output/{filename}.png"
                plt.savefig(path)
                plt.close()
                return path

            grafico_vend = salvar_grafico(resumo_vendedor, "Vendedor", "grafico_vendedor")
            grafico_est = salvar_grafico(resumo_estado, "Estado", "grafico_estado")
            grafico_prod = salvar_grafico(resumo_produto, "Produto", "grafico_produto")

            # --- PDF ---
            pdf_path = "output/relatorio.pdf"
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=14)
            pdf.cell(200, 10, txt="Relatório de Vendas - Diário", ln=True, align="C")
            pdf.ln(10)

            def adicionar_secao_pdf(titulo, resumo, imagem):
                pdf.set_font("Arial", size=12)
                pdf.cell(200, 10, txt=titulo, ln=True)
                for index, row in resumo.iterrows():
                    texto = f"{row[0]}: R$ {row['Valor Total']:.2f}"
                    pdf.cell(200, 10, txt=texto, ln=True)
                pdf.image(imagem, x=50, y=None, w=100)
                pdf.ln(10)

            adicionar_secao_pdf("🔹 Total Vendido por Vendedor", resumo_vendedor, grafico_vend)
            adicionar_secao_pdf("🔸 Total Vendido por Estado", resumo_estado, grafico_est)
            adicionar_secao_pdf("🛒 Total Vendido por Produto", resumo_produto, grafico_prod)

            pdf.output(pdf_path)

            # --- ENVIO POR E-MAIL ---
            msg = MIMEMultipart()
            msg["From"] = email_remetente
            msg["To"] = email_destino
            msg["Subject"] = "Relatório Diário de Vendas"
            corpo = "Olá,\n\nSegue em anexo o relatório diário de vendas com resumos por vendedor, estado e produto.\n\nAtenciosamente,\nSistema Automático"
            msg.attach(MIMEText(corpo, "plain"))
            with open(pdf_path, "rb") as file:
                part = MIMEApplication(file.read(), Name="relatorio.pdf")
                part["Content-Disposition"] = 'attachment; filename="relatorio.pdf"'
                msg.attach(part)

            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(email_remetente, senha_app)
            server.send_message(msg)
            server.quit()

            st.success("✅ Relatório enviado com sucesso!")

        except Exception as e:
            st.error(f"❌ Erro: {e}")
