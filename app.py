from flask import Flask, render_template, request, redirect, session, send_file
import pandas as pd
import io
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

app = Flask(__name__)
app.secret_key = "erp_secret"

@app.route('/generate_invoice', methods=['POST'])
def generate_invoice():
    styles = getSampleStyleSheet()

    customer = request.form['customer']
    part = request.form['part']
    qty = int(request.form['qty'])
    rate = float(request.form['rate'])

    total = qty * rate
    gst = total * 0.18
    final = total + gst

    file_path = "invoice.pdf"
    doc = SimpleDocTemplate(file_path)

    content = []

    # 🔥 COMPANY HEADER
    content.append(Paragraph("<b>RADIANCE POLYMERS</b>", styles['Title']))
    content.append(Paragraph("GSTIN: 27AAVFR6150R1Z4", styles['Normal']))
    content.append(Paragraph("State: Maharashtra", styles['Normal']))
    content.append(Spacer(1, 10))

    # 🔥 INVOICE DETAILS
    content.append(Paragraph("<b>TAX INVOICE</b>", styles['Heading2']))
    content.append(Paragraph(f"Invoice Date: 2026", styles['Normal']))
    content.append(Spacer(1, 10))

    # 🔥 BUYER DETAILS
    content.append(Paragraph(f"<b>Buyer:</b> {customer}", styles['Normal']))
    content.append(Spacer(1, 10))

    # 🔥 TABLE (LIKE YOUR IMAGE)
    data = [
        ["Sr No", "Description", "HSN", "Tax %", "Qty", "Rate", "Amount"],
        ["1", part, "87089900", "18%", qty, rate, total]
    ]

    table = Table(data, colWidths=[50, 150, 80, 60, 60, 60, 80])

    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.grey),
        ('TEXTCOLOR',(0,0),(-1,0),colors.white),
        ('GRID',(0,0),(-1,-1),1,colors.black),
        ('ALIGN',(0,0),(-1,-1),'CENTER')
    ]))

    content.append(table)
    content.append(Spacer(1, 20))

    # 🔥 TOTAL SECTION
    totals = [
        ["Subtotal", total],
        ["GST 18%", gst],
        ["Final Amount", final]
    ]

    total_table = Table(totals, colWidths=[200, 100])

    total_table.setStyle(TableStyle([
        ('GRID',(0,0),(-1,-1),1,colors.black)
    ]))

    content.append(total_table)

    doc.build(content)

    return send_file(file_path, as_attachment=True)



FILE = "parts.xlsx.xlsm"
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