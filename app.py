from datetime import datetime
import os

from flask import Flask, render_template, request, redirect, send_from_directory, session, send_file
import pandas as pd
import io
def get_next_invoice_no():
    import os
    import pandas as pd

    file = "invoices.xlsx"

    if not os.path.exists(file):
        return "INV-001" 
    
    df = pd.read_excel(file) 


    if len(df) == 0:
        return "INV-001"

    try:
        last = df.iloc[-1]["Invoice No"]
        num = int(last.split("-")[1])
        return f"INV-{num+1:03d}"
    except:
        return "INV-001"

app = Flask(__name__)
app.secret_key = "erp_secret"
FILE = "parts.xlsx.xlsm"

@app.route('/pdi')
def pdi():
    return render_template("pdi.html")
@app.route('/generate_pdi', methods=['POST'])
def generate_pdi():
    import pdfkit
    import os
    from datetime import datetime
    from flask import render_template, request, send_from_directory

    invoice_no = request.form['invoice_no']
    part = request.form['part']

    rows = [
        {"spec":"Dimension 14.8 ±0.20mm","inst":"PP/DVC"},
        {"spec":"Dimension 14.30 +0.20/-0.1 mm","inst":"PP/DVC"},
        {"spec":"Dimension 17.80 ±0.20mm","inst":"PP/DVC"},
        {"spec":"Dimension 13.50 ±0.30mm","inst":"PP/DVC"},
        {"spec":"Part Weight 3 ±0.5gm","inst":"WM"},
        {"spec":"Insertion force to Anchor 50N Max","inst":"UTM"},
        {"spec":"Removal force from Anchor 200N Min","inst":"UTM"}
    ]

    html = render_template(
        "pdi_template.html",
        invoice_no=invoice_no,
        part=part,
        date=datetime.now().strftime("%d-%m-%Y"),
        rows=rows
    )

    os.makedirs("static/pdi", exist_ok=True)

    file_path = f"static/pdi/PDI-{invoice_no}.pdf"

    config = pdfkit.configuration(
        wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
    )

    pdfkit.from_string(html, file_path, configuration=config)

    return send_from_directory("static/pdi", f"PDI-{invoice_no}.pdf", as_attachment=True)
@app.route('/generate_pdi', methods=['POST'])

def generate_pdi():

    import os
    from datetime import datetime
    from flask import send_from_directory
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet

    styles = getSampleStyleSheet()

    invoice_no = request.form['invoice_no']

    part = request.form['part']


    os.makedirs("static/pdi", exist_ok=True)

    file_name = f"static/pdi/PDI-{invoice_no}.pdf"

    doc = SimpleDocTemplate(file_name)
    content = []

    # 🔥 TITLE
    content.append(Paragraph(
        "<b>SUPPLIER P.D.I. REPORT CUM ITW RECEIVING INSPECTION REPORT</b>",
        styles['Title']
    ))
    content.append(Spacer(1, 10))

    # 🔥 TOP INFO TABLE
    info_data = [
        ["DATE", datetime.now().strftime("%d-%m-%Y"), "SUPPLIER", "Radiance Polymers", "Raw Material", "POM Celcon M90"],
        ["PART NAME", part, "PART NO", "62598", "Master Batch", ""],
        ["LOT QTY", "", "INVOICE NO", invoice_no, "Batch No", ""]
    ]

    info_table = Table(info_data, colWidths=[70, 120, 80, 150, 90, 120])
    info_table.setStyle(TableStyle([
        ('GRID',(0,0),(-1,-1),1,colors.black)
    ]))

    content.append(info_table)
    content.append(Spacer(1, 10))

    # 🔥 MAIN HEADER
    header = [
        "Sr No", "Specification", "Measuring Instrument",
        "CAV-1","CAV-2","CAV-3","CAV-4","CAV-5","CAV-6","CAV-7","CAV-8",
        "OK/NOK","1","2","3","4"
    ]

    table_data = [header]

    # 🔥 SAMPLE ROWS (like your Excel)
    specs = [
        "Dimension 14.8 ±0.20mm",
        "Dimension 14.30 +0.20/-0.1 mm",
        "Dimension 17.80 ±0.20mm",
        "Dimension 13.50 ±0.30mm",
        "Part Weight 3 ±0.5gm",
        "Insertion force to Anchor 50N Max",
        "Removal force from Anchor 200N Min"
    ]

    instruments = [
        "PP/DVC","PP/DVC","PP/DVC","PP/DVC",
        "WM","UTM","UTM"
    ]

    for i in range(len(specs)):
        row = [
            str(i+1),
            specs[i],
            instruments[i],
            "","","","","","","","","",
            "","","",""
        ]
        table_data.append(row)

    # 🔥 TABLE
    table = Table(table_data, repeatRows=1)

    table.setStyle(TableStyle([
        ('GRID',(0,0),(-1,-1),1,colors.black),
        ('BACKGROUND',(0,0),(-1,0),colors.grey),
        ('TEXTCOLOR',(0,0),(-1,0),colors.white),
        ('ALIGN',(0,0),(-1,-1),'CENTER'),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE')
    ]))

    content.append(table)

    doc.build(content)


    return send_from_directory("static/pdi", f"PDI-{invoice_no}.pdf", as_attachment=True)


@app.route('/generate_invoice', methods=['POST'])
def generate_invoice():
    import pandas as pd
    import os
    from datetime import datetime
    from flask import send_from_directory
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet

    styles = getSampleStyleSheet()

    customer = request.form['customer']
    parts = request.form.getlist('part[]')
    qtys = request.form.getlist('qty[]')
    rates = request.form.getlist('rate[]')


    # 🔥 AUTO INVOICE NUMBER
    invoice_no = get_next_invoice_no()

    # 🔥 ENSURE FOLDER EXISTS
    os.makedirs("static/invoices", exist_ok=True)

    # 🔥 FILE PATH (IMPORTANT FOR MOBILE)
    file_name = f"static/invoices/{invoice_no}.pdf"

    doc = SimpleDocTemplate(file_name)
    content = []

    # 🔥 HEADER (NO LOGO)
    content.append(Paragraph("<b>RADIANCE POLYMERS</b>", styles['Title']))
    content.append(Paragraph("GSTIN : 27AAVFR6150R1Z4 | State Code : 27 Maharashtra", styles['Normal']))
    content.append(Spacer(1, 10))

    # 🔥 INVOICE INFO
    info_data = [
        ["Invoice No:", invoice_no, "Date:", datetime.now().strftime("%d-%b-%Y")],
        ["P.O No:", "18040009076", "Payment Terms:", "30 Days"]
    ]

    info_table = Table(info_data, colWidths=[80, 150, 80, 150])
    info_table.setStyle(TableStyle([
        ('GRID',(0,0),(-1,-1),1,colors.black)
    ]))

    content.append(info_table)
    content.append(Spacer(1, 10))

    # 🔥 BUYER + CONSIGNEE
    buyer = [
        ["Buyer:", customer],
        ["Address:", "NANDUR, PUNE - 412202"],
        ["GSTIN:", "27AAACF3125C1Z9"],
        ["State:", "Maharashtra"]
    ]

    consignee = [
        ["Consignee:", customer],
        ["Address:", "NANDUR, PUNE - 412202"],
        ["GSTIN:", "27AAACF3125C1Z9"],
        ["State:", "Maharashtra"]
    ]

    buyer_table = Table(buyer)
    consignee_table = Table(consignee)

    buyer_table.setStyle(TableStyle([('GRID',(0,0),(-1,-1),1,colors.black)]))
    consignee_table.setStyle(TableStyle([('GRID',(0,0),(-1,-1),1,colors.black)]))

    content.append(Table([[buyer_table, consignee_table]]))
    content.append(Spacer(1, 10))

    # 🔥 MAIN TABLE
    table_data = [["Sr", "Description", "HSN/SAC", "Tax %", "Qty", "Rate", "Amount"]]

    total = 0

    for i in range(len(parts)):
        if parts[i]:
            q = float(qtys[i])
            r = float(rates[i])
            amt = q * r
            total += amt

            table_data.append([
                str(i+1),
                parts[i],
                "87089900",
                "18%",
                q,
                r,
                amt
            ])

    table = Table(table_data, colWidths=[40, 150, 80, 60, 60, 60, 80])

    table.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.grey),
        ('TEXTCOLOR',(0,0),(-1,0),colors.white),
        ('GRID',(0,0),(-1,-1),1,colors.black),
        ('ALIGN',(0,0),(-1,-1),'CENTER')
    ]))

    content.append(table)
    content.append(Spacer(1, 15))

    # 🔥 GST CALCULATION
    cgst = total * 0.09
    sgst = total * 0.09
    grand = total + cgst + sgst

    gst_table = [
        ["Subtotal", total],
        ["CGST 9%", cgst],
        ["SGST 9%", sgst],
        ["Grand Total", grand]
    ]

    t2 = Table(gst_table, colWidths=[200, 120])
    t2.setStyle(TableStyle([
        ('GRID',(0,0),(-1,-1),1,colors.black)
    ]))

    content.append(t2)
    content.append(Spacer(1, 20))

    # 🔥 SIGNATURE
    content.append(Paragraph("For Radiance Polymers", styles['Normal']))
    content.append(Spacer(1, 30))
    content.append(Paragraph("Authorized Signatory", styles['Normal']))

    # 🔥 BUILD PDF
    doc.build(content)

    # 🔥 SAVE HISTORY (EXCEL)
    file_path_excel = "invoices.xlsx"

    new_row = {
        "Invoice No": invoice_no,
        "Customer": customer,
        "Total": grand,
        "Date": datetime.now().strftime("%d-%m-%Y"),
        "File": f"{invoice_no}.pdf"
    }

    if os.path.exists(file_path_excel):
        df = pd.read_excel(file_path_excel)
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    else:
        df = pd.DataFrame([new_row])

    df.to_excel(file_path_excel, index=False)

    # 🔥 MOBILE DOWNLOAD FIX
    return send_from_directory(
        "static/invoices",
        f"{invoice_no}.pdf",
        as_attachment=True
    )


@app.route('/invoice')
def invoice():
    if 'user' not in session:
        return redirect('/login')
    return render_template("invoice.html")
@app.route('/invoice_history')
def invoice_history():
    if session.get('role') != "admin":
        return "Access Denied"

    df = pd.read_excel("invoices.xlsx")
    data = df.fillna("").values.tolist()

    return render_template("invoice_history.html", data=data)
@app.route('/download/<path:filename>')
def download(filename):
    return send_file(filename, as_attachment=True)
# ---------------- AUTO SHEET DETECT ----------------
def get_sheet():
    xls = pd.ExcelFile(FILE, engine="openpyxl")
    print("Available Sheets:", xls.sheet_names)
    return xls.sheet_names[0]


# ---------------- HOME ----------------
@app.route('/')
def home():
    return redirect('/login')


# ---------------- LOGIN ----------------
@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None

    if request.method == 'POST':
        u = request.form['username']
        p = request.form['password']

        if u == "admin" and p == "admin":
            session['user'] = "admin"
            session['role'] = "admin"
            return redirect('/dashboard')

        elif u == "user" and p == "user":
            session['user'] = "user"
            session['role'] = "user"
            return redirect('/inward')

        else:
            error = "Invalid Username or Password"

    return render_template("login.html", error=None)

# ---------------- LOGOUT ----------------
@app.route('/logout')
def logout():
    session.clear()
    return redirect('/login')


# ---------------- DASHBOARD ----------------
@app.route('/dashboard')
def dashboard():
    import pandas as pd

    try:
        df = pd.read_excel("invoices.xlsx")
    except:
        df = pd.DataFrame(columns=["Total"])

    total_sales = df['Total'].sum() if 'Total' in df else 0
    total_invoices = len(df)

    # 🔥 FIX: chart data
    chart_data = df['Total'].fillna(0).tolist() if 'Total' in df else []

    return render_template(
        "dashboard.html",
        total_sales=total_sales,
        total_invoices=total_invoices,
        chart_data=chart_data
    )
# ---------------- INWARD PAGE ----------------
@app.route('/inward')
def inward():
    if 'user' not in session:
        return redirect('/login')

    df = pd.read_excel(FILE, engine="openpyxl", sheet_name=get_sheet(), header=None)
    data = df.fillna("").values.tolist()

    return render_template("inward.html", data=data)


# ---------------- ADD INWARD ----------------
@app.route('/add_inward', methods=['POST'])
def add_inward():
    if 'user' not in session:
        return redirect('/login')

    df = pd.read_excel(FILE, engine="openpyxl", sheet_name=get_sheet(), header=None)

    new_data = [
        request.form.get('part_no'),
        request.form.get('description'),
        request.form.get('qty_in'),
        request.form.get('qty_out'),
        request.form.get('rejection'),
        session.get('user')   # 🔥 ADDED BY
    ]

    total_cols = df.shape[1]
    if len(new_data) < total_cols:
        new_data.extend([""] * (total_cols - len(new_data)))

    df.loc[len(df)] = new_data

    df.to_excel(FILE, index=False, header=False)

    return redirect('/inward')


# ---------------- STOCK ----------------
@app.route('/stock')
def stock():
    if session.get('role') != "admin":
        return "Access Denied (Admin Only)"

    df = pd.read_excel(FILE, engine="openpyxl", sheet_name=get_sheet(), header=None)

    part_no = request.args.get('part_no', '').strip()

    if part_no:
        df = df[df.astype(str).apply(
            lambda x: x.str.contains(part_no, case=False, na=False)
        ).any(axis=1)]

    data = df.fillna("").values.tolist()

    return render_template("stock.html", data=data)
@app.route('/activity')
def activity():
    if session.get('role') != "admin":
        return "Access Denied (Admin Only)"

    df = pd.read_excel(FILE, engine="openpyxl", sheet_name=get_sheet(), header=None)

    data = df.fillna("").values.tolist()

    return render_template("activity.html", data=data)
# ---------------- EXPORT ----------------
@app.route('/export')
def export():
    if session.get('role') != "admin":
        return "Access Denied (Admin Only)"

    df = pd.read_excel(FILE, engine="openpyxl", sheet_name=get_sheet(), header=None)

    part_no = request.args.get('part_no', '').strip()

    if part_no:
        df = df[df.astype(str).apply(
            lambda x: x.str.contains(part_no, case=False, na=False)
        ).any(axis=1)]

    output = io.BytesIO()
    df.to_excel(output, index=False, header=False)
    output.seek(0)

    return send_file(output, download_name="stock.xlsx", as_attachment=True)


# ---------------- RUN ----------------
if __name__ == "__main__":
    app.run(debug=True)