import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import os
from fpdf import FPDF
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText

# Page configuration
st.set_page_config(page_title="Sales Report", layout="centered")
st.title("📊 Sales Report Automation")
st.write("Upload an `.xlsx` spreadsheet to generate and send the report by email.")

# File upload
uploaded_file = st.file_uploader("Select the sales file (.xlsx)", type="xlsx")

# Email inputs
sender_email = st.text_input("Your email (Gmail)", "")
app_password = st.text_input("Gmail app password", type="password")
recipient_email = st.text_input("Recipient", "")

# When the button is clicked
if uploaded_file and sender_email and app_password and recipient_email:
    if st.button("📤 Generate Report and Send"):
        try:
            # Read the spreadsheet
            df = pd.read_excel(uploaded_file)

            # Create output folder
            os.makedirs("output", exist_ok=True)

            # --- SUMMARIES ---
            seller_summary = df.groupby("Seller")["Total Value"].sum().reset_index()
            state_summary = df.groupby("State")["Total Value"].sum().reset_index()
            product_summary = df.groupby("Product")["Total Value"].sum().reset_index()

          
            # --- PLOTS ---
            def save_plot(data, category, filename):
                plt.figure(figsize=(6, 4))
                plt.bar(data[category], data["Total Value"], color="mediumseagreen")
                plt.title(f"Total Sold by {category}")
                plt.ylabel("R$")
                plt.xticks(rotation=45)
                plt.tight_layout()
                path = f"output/{filename}.png"
                plt.savefig(path)
                plt.close()
                return path
                
            seller_plot = save_plot(seller_summary, "Seller", "seller_plot")
            state_plot = save_plot(state_summary, "State", "state_plot")
            product_plot = save_plot(product_summary, "Product", "product_plot")
            
            # --- PDF ---
            pdf_path = "output/report.pdf"
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=14)
            pdf.cell(200, 10, txt="Sales Report - Daily", ln=True, align="C")
            pdf.ln(10)

            def add_pdf_section(title, summary, image):
                pdf.set_font("Arial", size=12)
                pdf.cell(200, 10, txt=title, ln=True)
                for index, row in summary.iterrows():
                    text = f"{row[0]}: R$ {row['Valor Total']:.2f}"
                    pdf.cell(200, 10, txt=text, ln=True)
                pdf.image(image, x=50, y=None, w=100)
                pdf.ln(10)

            add_pdf_section("🔹 Total Sold by Seller", seller_summary, seller_plot)
            add_pdf_section("🔸 Total Sold by State", state_summary, state_plot)
            add_pdf_section("🛒 Total Sold by Product", product_summary, product_plot)

            pdf.output(pdf_path)

            # --- EMAIL SENDING ---
            msg = MIMEMultipart()
            msg["From"] = sender_email
            msg["To"] = recipient_email
            msg["Subject"] = "Daily Sales Report"
            body = "Hello,\n\nAttached is the daily sales report with summaries by seller, state, and product.\n\nBest regards,\nAutomated System"
            msg.attach(MIMEText(body, "plain"))
            with open(pdf_path, "rb") as file:
                part = MIMEApplication(file.read(), Name="report.pdf")
                part["Content-Disposition"] = 'attachment; filename="report.pdf"'
                msg.attach(part)

            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(sender_email, app_password)
            server.send_message(msg)
            server.quit()

            st.success("✅ Report sent successfully!")

        except Exception as e:
            st.error(f"❌ Error: {e}")
