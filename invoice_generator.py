import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Table, TableStyle
import os
from datetime import datetime

# --- SETTINGS ---
INPUT_FILE = "invoices.xlsx"
LOGO_PATH = "images/logo.jpg"
OUTPUT_DIR = "output"
FONT_PATH = "fonts/Poppins-Regular.ttf"
FONT_BOLD_PATH = "fonts/Poppins-Bold.ttf"  # If you have bold version

# Register fonts
pdfmetrics.registerFont(TTFont("Poppins", FONT_PATH))
if os.path.exists(FONT_BOLD_PATH):
    pdfmetrics.registerFont(TTFont("Poppins-Bold", FONT_BOLD_PATH))

os.makedirs(OUTPUT_DIR, exist_ok=True)

def clean_filename(text):
    """Make filename safe and readable"""
    return "".join(c if c.isalnum() or c in "_-" else "_" for c in text)

def create_invoice(invoice_no, data):
    client_name = data["Client Name"].iloc[0]
    invoice_date = data["Date"].iloc[0]

    # Convert date into readable format
    try:
        date_obj = pd.to_datetime(invoice_date)
        date_str = date_obj.strftime("%B %dth, %Y")  # e.g. October 20th, 2025
        file_date_str = date_obj.strftime("%d%b%Y")
    except:
        date_str = str(invoice_date)
        file_date_str = str(invoice_date).replace("/", "_")

    total = data["Total"].iloc[-1]

    # File name format
    safe_name = clean_filename(client_name)
    file_path = os.path.join(
        OUTPUT_DIR, f"{safe_name}_Invoice{int(invoice_no):02d}_{file_date_str}.pdf"
    )

    c = canvas.Canvas(file_path, pagesize=A4)
    width, height = A4

    # === Watermark (Logo as background) ===
    if os.path.exists(LOGO_PATH):
        c.saveState()
        # Position watermark in center, slightly lower
        watermark_size = 350
        c.translate(width / 2 - watermark_size / 2, height / 2 - watermark_size / 2 - 50)
        c.setFillAlpha(0.08)  # Very light watermark
        c.drawImage(LOGO_PATH, 0, 0, width=watermark_size, height=watermark_size, 
                   preserveAspectRatio=True, mask='auto')
        c.restoreState()

    # === Header Section ===
    header_y = height - 50
    
    # Logo at top left
    if os.path.exists(LOGO_PATH):
        logo_size = 55
        c.drawImage(LOGO_PATH, 30, header_y - logo_size, 
                   width=logo_size, height=logo_size, 
                   preserveAspectRatio=True, mask='auto')
    
    # Company name next to logo
    c.setFont("Poppins", 18)
    c.setFillColor(colors.black)
    c.drawString(95, header_y - 20, "Pixel")
    c.drawString(95, header_y - 40, "Pulse")
    
    # INVOICE title on right
    c.setFont("Poppins", 32)
    c.drawRightString(width - 30, header_y - 30, "INVOICE")
    
    # Horizontal line under header
    c.setStrokeColor(colors.black)
    c.setLineWidth(1)
    c.line(30, header_y - 70, width - 30, header_y - 70)

    # === Client Info Section ===
    info_y = header_y - 100
    c.setFont("Poppins", 10)
    c.drawString(30, info_y, "INVOICE TO :")
    
    c.setFont("Poppins", 16)
    c.drawString(30, info_y - 25, client_name)
    
    c.setFont("Poppins", 11)
    c.drawString(30, info_y - 50, f"Invoice No : {int(invoice_no):02d}")
    c.drawString(30, info_y - 68, f"Invoice Date : {date_str}")

    # === Table Section ===
    table_y = info_y - 120
    
    # Table headers with background
    c.setFillColor(colors.HexColor("#F5F5F5"))
    c.rect(30, table_y - 5, width - 60, 25, fill=1, stroke=0)
    
    c.setFillColor(colors.black)
    c.setFont("Poppins", 11)
    c.drawString(50, table_y + 5, "NAME")
    c.drawString(320, table_y + 5, "QTY")
    c.drawString(410, table_y + 5, "PRICE")
    c.drawRightString(width - 50, table_y + 5, "TOTAL")
    
    # Table line
    c.setStrokeColor(colors.black)
    c.setLineWidth(0.5)
    c.line(30, table_y - 5, width - 30, table_y - 5)
    
    # Table rows
    row_y = table_y - 30
    c.setFont("Poppins", 10)
    
    for idx, row in data.iterrows():
        c.drawString(50, row_y, str(row["Service"]))
        c.drawString(330, row_y, f"{int(row['Qty']):02d}")
        c.drawString(410, row_y, f"PKR {int(row['Amount']):,}")
        c.drawRightString(width - 50, row_y, f"PKR {int(row['Subtotal']):,}")
        
        # Light separator line
        row_y -= 20
        c.setStrokeColor(colors.HexColor("#E0E0E0"))
        c.setLineWidth(0.3)
        c.line(30, row_y + 5, width - 30, row_y + 5)
        row_y -= 10

    # === Totals Section ===
    totals_y = row_y - 30
    c.setFont("Poppins", 11)
    c.drawRightString(width - 180, totals_y, "Sub-total :")
    c.drawRightString(width - 50, totals_y, f"PKR {int(total):,}")
    
    totals_y -= 25
    c.setFont("Poppins", 13)
    c.drawRightString(width - 180, totals_y, "Total :")
    c.drawRightString(width - 50, totals_y, f"PKR {int(total):,}")

    # === Bottom Section - Bank Details & Terms ===
    bottom_y = 220
    
    # Bank Details
    c.setFont("Poppins", 11)
    c.drawString(30, bottom_y, "BANK DETAILS")
    
    c.setFont("Poppins", 9)
    c.drawString(30, bottom_y - 18, "MEEZAN BANK")
    c.drawString(30, bottom_y - 33, "SYED MUHAMMAD HASAN JAWAID")
    c.drawString(30, bottom_y - 48, "PK32MEZN0001930109968206")
    
    c.drawString(30, bottom_y - 70, "SADAPAY")
    c.drawString(30, bottom_y - 85, "HASAN JAWAID")
    c.drawString(30, bottom_y - 100, "0310 246134")
    
    # Terms
    c.setFont("Poppins", 11)
    c.drawString(330, bottom_y, "TERMS")
    
    c.setFont("Poppins", 8)
    c.drawString(330, bottom_y - 18, "PAYMENT IS DUE WITHIN 3 DAYS FROM")
    c.drawString(330, bottom_y - 30, "THE DATE OF INVOICE SENT.")
    c.drawString(330, bottom_y - 42, "MAKE PAYMENT TO THE MENTIONED")
    c.drawString(330, bottom_y - 54, "ACCOUNTS ONLY. IN CASE OF")
    c.drawString(330, bottom_y - 66, "QUESTIONS PLEASE REACH OUT TO US.")

    # === Footer Section ===
    # Top line
    c.setStrokeColor(colors.black)
    c.setLineWidth(1)
    c.line(30, 80, width - 30, 80)
    
    # Phone icon (simple drawing)
    c.setFillColor(colors.HexColor("#FF6B6B"))
    c.setStrokeColor(colors.HexColor("#FF6B6B"))
    c.setLineWidth(2)
    c.roundRect(33, 52, 12, 16, 2, fill=0, stroke=1)
    c.setLineWidth(1)
    c.line(35, 65, 43, 65)
    c.line(35, 54, 43, 54)
    
    c.setFillColor(colors.black)
    c.setFont("Poppins", 9)
    c.drawString(50, 55, "0310-2461341")
    c.drawString(50, 40, "0309-2511984")
    
    # Email icon (simple envelope)
    c.setFillColor(colors.HexColor("#4A90E2"))
    c.setStrokeColor(colors.HexColor("#4A90E2"))
    c.setLineWidth(1.5)
    c.rect(203, 52, 16, 12, fill=0, stroke=1)
    c.line(203, 64, 211, 58)
    c.line(219, 64, 211, 58)
    
    c.setFillColor(colors.black)
    c.setFont("Poppins", 9)
    c.drawString(224, 55, "hasan92jawed@gmail.com")
    c.drawString(224, 40, "shaikhumair297@gmail.com")
    
    # Thank you
    c.setFont("Poppins", 13)
    c.drawRightString(width - 30, 50, "THANK YOU!")

    c.save()
    print(f"✅ Created {file_path}")


def main():
    if not os.path.exists(INPUT_FILE):
        print(f"❌ Error: {INPUT_FILE} not found!")
        return
    
    if not os.path.exists(LOGO_PATH):
        print(f"⚠️  Warning: {LOGO_PATH} not found. Invoice will be created without logo.")
    
    df = pd.read_excel(INPUT_FILE)
    
    # Group by invoice number and create PDFs
    for invoice_no, group in df.groupby("Invoice No."):
        create_invoice(invoice_no, group)
    
    print(f"\n✨ All invoices generated successfully in '{OUTPUT_DIR}' folder!")


if __name__ == "__main__":
    main()