from flask import Flask, config, render_template, request, send_from_directory
from datetime import date, datetime
import os

from flask import Flask, render_template, request, redirect, send_from_directory, session, send_file
import pandas as pd
import io
import datetime

def get_next_invoice_no():
    return "INV-" + datetime.datetime.now().strftime("%Y%m%d%H%M%S")
from sqlalchemy import values
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
    import pandas as pd

    try:
        df = pd.read_excel("invoices.xlsx")

        # 🔥 DEBUG PRINT (IMPORTANT)
        print(df.columns)

        invoices = df.to_dict(orient="records")

    except Exception as e:
        print("ERROR:", e)
        invoices = []

    specs = [
        "5.10 ±0.05 mm",
        "6.05 ±0.1 mm",
        "16.1 ±0.2 mm",
        "7.5 ±0.15 mm",
        "15.3 ±0.1 mm",
        "Insertion force 50N",
        "Removal force 200N"
    ]

    return render_template("pdi.html", invoices=invoices, specs=specs)
def check_ok_nok(value, spec):
    try:
        if "±" in spec:
            base, tol = spec.split("±")
            base = float(base.strip())
            tol = float(tol.strip().split()[0])

            min_val = base - tol
            max_val = base + tol

            val = float(value)

            return "OK" if min_val <= val <= max_val else "NOK"

        else:
            return "OK"  # for non-numeric specs

    except:
        return "OK"
    
@app.route('/generate_pdi', methods=['POST'])
def generate_pdi():
    import pdfkit, os
    from datetime import datetime
    from flask import render_template, request, send_from_directory

    # 🔥 FORM DATA (SAFE)
    company = request.form.get('company', '')
    invoice_no = request.form.get('invoice_no', '')
    part = request.form.get('part', '')
    customer = request.form.get('customer', '')

    # 🔥 SPEC DATA
    rows = [
        {"spec":"5.10 ±0.05","inst":"PP/DVC"},
        {"spec":"6.05 ±0.10","inst":"PP/DVC"},
        {"spec":"16.10 ±0.20","inst":"PP/DVC"},
        {"spec":"7.50 ±0.15","inst":"PP/DVC"},
        {"spec":"15.30 ±0.10","inst":"PP/DVC"},
        {"spec":"50","inst":"UTM"},
        {"spec":"200","inst":"UTM"}
    ]

    # 🔥 AESTHETIC
    aesthetic = [
        "Colour OK",
        "No scratches",
        "No damage",
        "Proper labeling"
    ]

    # 🔥 SIMPLE SAFE CHECK FUNCTION
    def safe_check(val, spec):
        try:
            if "±" in spec:
                base, tol = spec.split("±")
                base = float(base.strip())
                tol = float(tol.strip())

                min_v = base - tol
                max_v = base + tol

                v = float(val)

                if min_v <= v <= max_v:
                    return "OK"
                else:
                    return "NOK"
            else:
                return "OK"
        except:
            return "OK"

    # 🔥 VALUES (NO ERROR VERSION)
    values = []

    for i in range(7):
        row_vals = []
        status = "OK"

        for j in range(6):
            val = request.form.get(f"val{i}_{j}", "")
            row_vals.append(val)

            if val != "":
                if safe_check(val, rows[i]["spec"]) == "NOK":
                    status = "NOK"

        values.append({
            "vals": row_vals,
            "ok": status
        })

    # 🔥 TEMPLATE SELECT
    if company == "fleetguard":
        template = "pdi_fleetguard.html"
    elif company == "kinetic":
        template = "pdi_kinetic.html"
    else:
        template = "pdi_template.html"

    # 🔥 RENDER
    html = render_template(
        template,
        invoice_no=invoice_no,
        part=part,
        customer=customer,
        date=datetime.now().strftime("%d-%m-%Y"),
        rows=rows,
        aesthetic=aesthetic,
        values=values
    )

    # 🔥 SAVE PDF
    os.makedirs("static/pdi", exist_ok=True)

    file_path = f"static/pdi/PDI-{invoice_no}.pdf"


    config = pdfkit.configuration(
        wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
    )

    options = {
        'orientation': 'Landscape',
        'page-size': 'A4',
        'margin-top': '5mm',
        'margin-bottom': '5mm',
        'margin-left': '5mm',
        'margin-right': '5mm'
    }

    pdfkit.from_string(html, file_path, configuration=config, options=options)


    return send_from_directory("static/pdi", f"PDI-{invoice_no}.pdf", as_attachment=True)


@app.route('/generate_invoice', methods=['POST'])
def generate_invoice():
    import pdfkit, os
    from flask import request, render_template, send_from_directory

    invoice_no = get_next_invoice_no()

    customer = request.form.get('customer', '')
    date = request.form.get('date', '')
    items = request.form.getlist('item[]')
    qtys = request.form.getlist('qty[]')
    prices = request.form.getlist('price[]')

    data = []
    total = 0

    for i in range(len(items)):
        try:
            qty = float(qtys[i])
            price = float(prices[i])
            amount = qty * price
        except:
            qty = 0
            price = 0
            amount = 0

        total += amount

        data.append({
            "item": items[i],
            "qty": qty,
            "price": price,
            "amount": amount
        })

    html = render_template(
        "invoice.html",
        invoice_no=invoice_no,
        customer=customer,
        date=date,
        data=data,
        total=total
    )

    os.makedirs("static/invoices", exist_ok=True)
    file_path = f"static/invoices/{invoice_no}.pdf"

    config = pdfkit.configuration(
        wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
    )

    pdfkit.from_string(html, file_path, configuration=config)

    return send_from_directory("static/invoices", f"{invoice_no}.pdf", as_attachment=True)
@app.route('/invoice')
def invoice_page():
    return render_template("invoice_form.html")

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
    import os

    invoice_folder = "static/invoices"
    pdi_folder = "static/pdi"

    invoices = []
    pdis = []

    # 🔥 SAFE FILE READ
    if os.path.exists(invoice_folder):
        invoices = os.listdir(invoice_folder)

    if os.path.exists(pdi_folder):
        pdis = os.listdir(pdi_folder)

    # 🔥 TOTALS
    total_invoices = len(invoices)
    total_pdi = len(pdis)

    # 🔥 DUMMY SALES (SAFE)
    total_sales = total_invoices * 1000   # you can improve later

    # 🔥 CHART DATA (ALWAYS LIST)
    chart_data = [total_invoices, total_pdi]

    return render_template(
        "dashboard.html",
        total=total_invoices + total_pdi,
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
import json
from flask import request

@app.route('/save_inward', methods=['POST'])
def save_inward():
    sr = request.form.getlist('sr[]')
    desc = request.form.getlist('desc[]')
    part_no = request.form.getlist('part_no[]')
    rate = request.form.getlist('rate[]')
    prod = request.form.getlist('prod_qty[]')
    wip = request.form.getlist('wip_qty[]')
    fg = request.form.getlist('fg_qty[]')
    total = request.form.getlist('total[]')

    data = []

    for i in range(len(sr)):
        if part_no[i]:  # skip empty rows
            data.append({
                "sr": sr[i],
                "desc": desc[i],
                "part_no": part_no[i],
                "rate": rate[i],
                "prod": prod[i],
                "wip": wip[i],
                "fg": fg[i],
                "total": total[i]
            })

    # 🔥 SAVE TO FILE
    try:
        with open("inward_data.json", "r") as f:
            old_data = json.load(f)
    except:
        old_data = []

    old_data.extend(data)

    with open("inward_data.json", "w") as f:
        json.dump(old_data, f, indent=4)

    return "Inward Saved Successfully ✅"


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
    import pandas as pd
    import json
    from flask import request

    data = []

    # 🔥 STEP 1: READ EXCEL (SIMPLE)
    try:
        df = pd.read_excel("parts.xlsx.xlsm", engine="openpyxl", sheet_name=0)

        print("✅ Excel Loaded")
        print(df.head())

        # clean
        df = df.dropna(how='all')
        df = df.fillna("")

        # 🔥 DIRECT COLUMN ACCESS (NO AUTO DETECT NOW)
        for _, row in df.iterrows():
            data.append({
                "sr": row.iloc[0],
                "desc": row.iloc[1],
                "part_no": str(row.iloc[2]),
                "rate": row.iloc[3],
                "prod": row.iloc[4],
                "wip": row.iloc[5],
                "fg": row.iloc[6],
                "total": row.iloc[7]
            })

    except Exception as e:
        print("❌ Excel error:", e)

    # 🔥 STEP 2: ADD INWARD DATA
    try:
        with open("inward_data.json", "r") as f:
            inward = json.load(f)
            data.extend(inward)
    except:
        pass

    # 🔍 SEARCH
    part_no = request.args.get("part_no")
    if part_no:
        data = [d for d in data if part_no.lower() in str(d["part_no"]).lower()]

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