from flask import Flask, render_template, request, redirect, session, send_file
import pandas as pd
import io
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from datetime import datetime 

app = Flask(__name__)
app.secret_key = "erp_secret"
FILE = "parts.xlsx.xlsm"


@app.route('/generate_invoice', methods=['POST'])
def generate_invoice():
    styles = getSampleStyleSheet()

    customer = request.form['customer']
    parts = request.form.getlist('part[]')
    qtys = request.form.getlist('qty[]')
    rates = request.form.getlist('rate[]')

    doc = SimpleDocTemplate("invoice.pdf")
    content = []

    # 🔥 HEADER
    content.append(Paragraph("<b>RADIANCE POLYMERS</b>", styles['Title']))
    content.append(Paragraph("GSTIN : 27AAVFR6150R1Z4 | State Code : 27 Maharashtra", styles['Normal']))
    content.append(Spacer(1, 10))

    # 🔥 INVOICE INFO (TOP BOX LIKE IMAGE)
    info_data = [
        ["Invoice No:", "RPG/B/0010/26-27", "Date:", datetime.now().strftime("%d-%b-%Y")],
        ["P.O No:", "18040009076", "Payment Terms:", "30 Days"]
    ]

    info_table = Table(info_data, colWidths=[80, 150, 80, 150])
    info_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 1, colors.black)
    ]))
    content.append(info_table)
    content.append(Spacer(1, 10))

    # 🔥 BUYER + CONSIGNEE SIDE BY SIDE
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

    side_by_side = Table([[buyer_table, consignee_table]])
    content.append(side_by_side)
    content.append(Spacer(1, 10))

    # 🔥 MAIN ITEM TABLE (MULTI ROW)
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

    # 🔥 GST + TOTAL SECTION (LIKE IMAGE RIGHT SIDE)
    cgst = total * 0.09
    sgst = total * 0.09
    grand = total + cgst + sgst

    totals = [
        ["Subtotal", total],
        ["CGST 9%", cgst],
        ["SGST 9%", sgst],
        ["Grand Total", grand]
    ]

    total_table = Table(totals, colWidths=[200, 120])
    total_table.setStyle(TableStyle([
        ('GRID',(0,0),(-1,-1),1,colors.black)
    ]))

    content.append(total_table)
    content.append(Spacer(1, 20))

    # 🔥 FOOTER (SIGN)
    content.append(Paragraph("For Radiance Polymers", styles['Normal']))
    content.append(Spacer(1, 30))
    content.append(Paragraph("Authorized Signatory", styles['Normal']))

    doc.build(content)

    return send_file("invoice.pdf", as_attachment=True)
@app.route('/invoice')
def invoice():
    if 'user' not in session:
        return redirect('/login')
    return render_template("invoice.html")

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
    if session.get('role') != "admin":
        return "Access Denied (Admin Only)"

    df = pd.read_excel(FILE, engine="openpyxl", sheet_name=get_sheet(), header=None)
    total = len(df)
    col_counts = df.count().tolist() if not df.empty else []

    return render_template("dashboard.html", total=total, chart_data=col_counts)
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