import os

from flask import Flask, render_template, request, redirect, session, send_file
import pandas as pd
import io


app = Flask(__name__)
app.secret_key = "erp_secret"
FILE = "parts.xlsx.xlsm"


@app.route('/generate_invoice', methods=['POST'])
def generate_invoice():
    import pandas as pd
    import os
    from datetime import datetime
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet

    styles = getSampleStyleSheet()

    customer = request.form['customer']
    parts = request.form.getlist('part[]')
    qtys = request.form.getlist('qty[]')
    rates = request.form.getlist('rate[]')

    # 🔥 CREATE UNIQUE INVOICE NO
    invoice_no = "INV-" + datetime.now().strftime("%Y%m%d%H%M%S")
    file_name = f"{invoice_no}.pdf"

    doc = SimpleDocTemplate(file_name)
    content = []

    # HEADER
    content.append(Paragraph("<b>RADIANCE POLYMERS</b>", styles['Title']))

    content.append(Spacer(1, 10))

    # TABLE
    table_data = [["Sr", "Description", "Qty", "Rate", "Amount"]]

    total = 0

    for i in range(len(parts)):
        if parts[i]:
            q = float(qtys[i])
            r = float(rates[i])
            amt = q * r
            total += amt

            table_data.append([i+1, parts[i], q, r, amt])

    table = Table(table_data)
    table.setStyle(TableStyle([
        ('GRID',(0,0),(-1,-1),1,colors.black)
    ]))

    content.append(table)

    # GST
    cgst = total * 0.09
    sgst = total * 0.09
    grand = total + cgst + sgst

    content.append(Spacer(1, 10))
    content.append(Paragraph(f"Total: {total}", styles['Normal']))
    content.append(Paragraph(f"CGST: {cgst}", styles['Normal']))
    content.append(Paragraph(f"SGST: {sgst}", styles['Normal']))
    content.append(Paragraph(f"<b>Grand Total: {grand}</b>", styles['Normal']))

    # BUILD PDF
    doc.build(content)

    # 🔥 SAVE HISTORY
    file_path_excel = "invoices.xlsx"

    new_row = {
        "Invoice No": invoice_no,
        "Customer": customer,
        "Total": grand,
        "Date": datetime.now().strftime("%d-%m-%Y"),
        "File": file_name
    }

    if os.path.exists(file_path_excel):
        df = pd.read_excel(file_path_excel)
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    else:
        df = pd.DataFrame([new_row])

    df.to_excel(file_path_excel, index=False)

    return send_file(file_name, as_attachment=True)
   
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
@app.route('/download/<filename>')
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