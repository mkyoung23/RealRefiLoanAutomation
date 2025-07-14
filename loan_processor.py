import os
import pandas as pd
import numpy as np
from datetime import datetime
import logging
from openai import OpenAI
import re
import time
from dotenv import load_dotenv
import streamlit as st
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from io import BytesIO
from functools import lru_cache
import json
from docx import Document
from docx.shared import Inches, RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Load environment variables
load_dotenv()
OPENAI_API_KEY = os.getenv('OPENAI_API_KEY', '')

OUTPUT_FOLDER = "BorrowerResults"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

RATE_MATRIX = [
    (0.0, 0.6, 0.0575, 0.06, 0.06125),
    (0.6001, 0.7, 0.05875, 0.06125, 0.0625),
    (0.7001, 0.75, 0.06, 0.0625, 0.06375),
    (0.7501, 0.8, 0.06125, 0.06375, 0.065),
    (0.8001, 0.95, 0.06125, 0.06375, 0.065)
]

NO_COST_RATE_ADJ = 0.0025
ROLL_RATE_ADJ = -0.0025
CASHOUT_RATE_ADJ = 0.00125

APPRECIATION = 1.07
MAX_CASHOUT_LTV = 0.80
PMT_TOLERANCE = 5.0  # Dollars for PMT validation

# Logging
logging.basicConfig(filename=os.path.join(OUTPUT_FOLDER, 'loan_processing.log'), level=logging.DEBUG, format='%(asctime)s [%(levelname)s] %(message)s')

def clean_currency(value):
    if pd.isna(value): return 0.0
    if isinstance(value, str):
        return float(value.replace('$', '').replace(',', '').strip())
    return float(value)

def clean_percentage(value):
    if pd.isna(value): return 0.0
    if isinstance(value, str):
        value = value.replace('%', '').strip()
    try:
        val = float(value)
    except:
        return 0.0
    return val / 100 if val > 1 else val

def get_rates_from_ltv(ltv):
    for lower, upper, rate_15, rate_20, rate_30 in RATE_MATRIX:
        if lower <= ltv <= upper:
            return rate_15, rate_20, rate_30
    return 0.06125, 0.06375, 0.065

def calculate_pmt(rate, nper, pv):
    if rate == 0 or pv <= 0 or nper <= 0:
        return 0.0
    try:
        res = -np.pmt(rate, nper, pv)
        logging.debug(f"PMT: rate={rate}, nper={nper}, pv={pv}, result={res}")
        return res
    except Exception as e:
        logging.error(f"PMT error: {e}")
        return 0.0

def calculate_amortized_balance(principal, rate, nper, payments_made):
    if principal <= 0 or payments_made <= 0 or nper <= 0:
        return principal
    try:
        remaining_nper = nper - payments_made
        pmt = calculate_pmt(rate, nper, principal)
        balance = np.pv(rate, remaining_nper, -pmt)
        logging.debug(f"AMORT BAL: principal={principal}, rate={rate}, nper={nper}, payments_made={payments_made}, result={balance}")
        return max(0, balance)
    except Exception as e:
        logging.error(f"AMORT error: {e}")
        return principal

def validate_pmt(row):
    calc_pmt = calculate_pmt(row['Current Interest Rate']/12, row['Loan Term (years)']*12, row['Total Original Loan Amount'])
    if abs(calc_pmt - row['Current P&I Mtg Pymt']) > PMT_TOLERANCE:
        logging.warning(f"PMT mismatch for {row['Borrower First Name']}: Calc {calc_pmt:.2f} vs Input {row['Current P&I Mtg Pymt']:.2f}")
        return "Warning: PMT Mismatch"
    return "Valid"

def calculate_months_elapsed(first_payment_date):
    if pd.isna(first_payment_date):
        return 0
    try:
        first_payment = pd.to_datetime(first_payment_date)
        today = datetime.now()
        months = (today.year - first_payment.year) * 12 + (today.month - first_payment.month)
        return max(months, 0)
    except:
        return 0

@lru_cache(maxsize=128)
def estimate_home_value(address, original_value, months_elapsed):
    if not OPENAI_API_KEY:
        return original_value * (APPRECIATION ** (months_elapsed / 12))
    try:
        client = OpenAI(api_key=OPENAI_API_KEY)
        prompt = f"Estimate current market value for {address}, NJ, based on recent comps, trends, and 7% YoY appreciation from {original_value}. Return only the number."
        response = client.chat.completions.create(model="gpt-4o-mini", messages=[{"role": "user", "content": prompt}], max_tokens=20)
        content = response.choices[0].message.content.strip()
        match = re.search(r'\d+', content.replace(',', ''))
        return float(match.group()) if match else original_value * (APPRECIATION ** (months_elapsed / 12))
    except Exception as e:
        logging.error(f"Home value estimation error: {e}")
        return original_value * (APPRECIATION ** (months_elapsed / 12))

def batch_generate_texts(df, officer_name, company_name, app_link):
    if not OPENAI_API_KEY:
        return [{} for _ in range(len(df))]
    try:
        client = OpenAI(api_key=OPENAI_API_KEY)
        prompts = [f"Generate 10 urgent, personalized texts (5 types x 2 A/B variants) for refi outreach based on {row.to_dict()}. Make them sound like from the borrower's previous loan officer ({officer_name} at {company_name}). Highlight current payment vs new options (regular, no-cost, roll-in, cash-out, HELOC), savings, equity, market trends, FOMO (e.g., 'Rates rising soon!'). Include CTA with {app_link}. Vary for response testing." for _, row in df.iterrows()]
        responses = []
        for p in prompts:
            response = client.chat.completions.create(model="gpt-4o-mini", messages=[{"role": "user", "content": p}], max_tokens=600)
            responses.append(response)
            time.sleep(1)  # Rate limit
        texts_list = []
        for resp in responses:
            content = resp.choices[0].message.content
            texts = {line.split(':')[0].strip(): ':'.join(line.split(':')[1:]).strip() for line in content.split('\n') if ':' in line}
            texts_list.append(texts)
        return texts_list
    except Exception as e:
        logging.error(f"Text generation error: {e}")
        return [{} for _ in range(len(df))]

def generate_savings_chart(borrower_data):
    fig, ax = plt.subplots()
    labels = ['Current PMT', 'New Reg PMT', 'Monthly Savings']
    values = [borrower_data['Current P&I Mtg Pymt'], borrower_data['Pmt Reg (30yr)'], borrower_data['Savings Reg']]
    colors = ['red' if v < 0 else 'green' for v in values]
    ax.bar(labels, values, color=colors)
    ax.set_title('Payment & Savings Breakdown')
    ax.set_ylabel('$')
    buf = BytesIO()
    plt.savefig(buf, format='png')
    buf.seek(0)
    return buf

def generate_word_report(borrower_data, texts, officer_name, company_name, chart_buf):
    doc = Document()
    doc.add_heading(f"Personalized Refinance Report for {borrower_data['Borrower First Name']} {borrower_data['Borrower Last Name']}", 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Prepared by {officer_name}, {company_name} | Date: {datetime.now().strftime('%Y-%m-%d')}").style.font.color.rgb = RGBColor(0, 0, 0)

    # Executive Summary
    summary = doc.add_paragraph()
    summary.add_run("Executive Summary").bold = True
    summary.add_run(f"\nCurrent Payment: ${borrower_data['Current P&I Mtg Pymt']:.2f}/mo\nEstimated Home Value: ${borrower_data['New Estimated Home Value']:.2f} (up ${borrower_data['Equity Increase ($)']:.2f})\nTop Opportunity: Save up to ${max(borrower_data['Savings Reg'], borrower_data['Savings Roll']):.2f}/mo with roll-in option.")
    summary.style.font.size = Pt(12)
    summary.style.font.color.rgb = RGBColor(0, 128, 0) if borrower_data['Savings Reg'] > 0 else RGBColor(255, 0, 0)

    # Payment Comparisons Table
    doc.add_heading("Payment Comparisons", level=1)
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Option'
    hdr_cells[1].text = 'New Rate'
    hdr_cells[2].text = 'New Payment'
    hdr_cells[3].text = 'Monthly Savings'
    hdr_cells[4].text = 'Details'
    options = [
        ('Regular Refi (30yr)', borrower_data['New Rate (30yr)'], borrower_data['Pmt Reg (30yr)'], borrower_data['Savings Reg'], 'Standard refinance'),
        ('No-Cost Refi (30yr)', borrower_data['NoCost Rate (30yr)'], borrower_data['Pmt NoCost (30yr)'], borrower_data['Savings NoCost'], '$0 closing costs'),
        ('Roll-In Refi (30yr)', borrower_data['Roll Rate (30yr)'], borrower_data['Pmt Roll (30yr)'], borrower_data['Savings Roll'], 'Roll points into loan'),
        ('Cash-Out Refi (30yr)', borrower_data['CashOut Rate (30yr)'], borrower_data['Pmt CashOut (30yr)'], borrower_data['Savings CashOut'], f"Cash out up to ${borrower_data['Max CashOut Amount']:.2f}"),
        ('HELOC Option', 'Variable', 'Custom', 'Varies', f"Tap ${borrower_data['Equity Increase ($)']:.2f} equity at low rates")
    ]
    for opt, rate, pmt, sav, det in options:
        row_cells = table.add_row().cells
        row_cells[0].text = opt
        row_cells[1].text = f"{rate:.3%}" if isinstance(rate, float) else rate
        row_cells[2].text = f"${pmt:.2f}" if isinstance(pmt, float) else pmt
        row_cells[3].text = f"${sav:.2f}" if isinstance(sav, float) else sav
        row_cells[3].paragraphs[0].runs[0].font.color.rgb = RGBColor(0, 128, 0) if (isinstance(sav, float) and sav > 0) else RGBColor(255, 0, 0)
        row_cells[4].text = det

    # Embed Chart
    doc.add_heading("Visual Savings Breakdown", level=1)
    doc.add_picture(chart_buf, width=Inches(5.5))

    # Personalized Texts
    doc.add_heading("Copy-Paste Outreach Texts", level=1)
    for key, text in texts.items():
        p = doc.add_paragraph()
        p.add_run(f"{key}: ").bold = True
        p.add_run(text)

    # Custom Notes Section
    doc.add_heading("Custom Notes from Officer", level=1)
    doc.add_paragraph("Add your personalized notes here for this borrower.")

    doc.save(os.path.join(OUTPUT_FOLDER, f"{borrower_data['Borrower First Name']}_{borrower_data['Borrower Last Name']}_Report.docx"))
    return os.path.join(OUTPUT_FOLDER, f"{borrower_data['Borrower First Name']}_{borrower_data['Borrower Last Name']}_Report.docx")

def generate_pdf_backup(borrower_data, texts, officer_name, company_name, chart_buf):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    # Similar to previous PDF generation for backup
    c.save()
    buf.seek(0)
    with open(os.path.join(OUTPUT_FOLDER, f"{borrower_data['Borrower First Name']}_backup_report.pdf"), 'wb') as f:
        f.write(buf.getvalue())

def generate_email_template(borrower_data, texts, officer_name, company_name):
    name = borrower_data['Borrower First Name']
    template = f"Subject: Hey {name}, Remember Your Loan with Me at {company_name}? Big Savings Opportunity Inside!\n\nHi {name},\nIt's {officer_name} from {company_name}—we worked on your loan back in {borrower_data['First Pymt Date'].year}. With rates changing, here's how you can save compared to your current ${borrower_data['Current P&I Mtg Pymt']:.2f}/mo:\n\n"
    for key, text in texts.items():
        template += f"{key}:\n{text}\n\n"
    template += "Rates are rising soon—reply or click {app_link} to lock in! Best,\n{officer_name}"
    return template

def simulate_scenario(balance, home_value, custom_rate, custom_term, custom_cashout=0):
    monthly_rate = custom_rate / 12
    nper = custom_term * 12
    loan_amount = balance + custom_cashout
    pmt = calculate_pmt(monthly_rate, nper, loan_amount)
    ltv = loan_amount / home_value if home_value > 0 else 0
    return pmt, ltv

def process_loans(df, officer_name, company_name, app_link, progress_bar):
    progress_bar.progress(0)
    # Cleaning
    df['Total Original Loan Amount'] = df['Total Original Loan Amount'].apply(clean_currency)
    df['Original Appraised Value'] = df['Original Appraised Value'].apply(clean_currency)
    df['Current Interest Rate'] = df['Current Interest Rate'].apply(clean_percentage)
    df['Current P&I Mtg Pymt'] = df['Current P&I Mtg Pymt'].apply(clean_currency)
    df['Loan Term (years)'] = pd.to_numeric(df['Loan Term (years)'], errors='coerce').fillna(30)
    df['First Pymt Date'] = pd.to_datetime(df['First Pymt Date'], errors='coerce')

    progress_bar.progress(0.1)
    df['Months Elapsed'] = df['First Pymt Date'].apply(calculate_months_elapsed)

    progress_bar.progress(0.2)
    for idx, row in df.iterrows():
        address = f"{row['Subject Property Address']}, {row['Subject Property City']}, {row['Subject Property State']}"
        est_value = estimate_home_value(address, row['Original Appraised Value'], row['Months Elapsed'])
        df.at[idx, 'New Estimated Home Value'] = est_value

    progress_bar.progress(0.3)
    df['New Loan Balance'] = df.apply(lambda row: calculate_amortized_balance(
        row['Total Original Loan Amount'], row['Current Interest Rate']/12, row['Loan Term (years)']*12, row['Months Elapsed']
    ), axis=1)

    df['LTV'] = df['New Loan Balance'] / df['New Estimated Home Value']

    progress_bar.progress(0.4)
    rate_data = df['LTV'].apply(get_rates_from_ltv)
    df['New Rate (15yr)'] = rate_data.apply(lambda x: x[0])
    df['New Rate (20yr)'] = rate_data.apply(lambda x: x[1])
    df['New Rate (30yr)'] = rate_data.apply(lambda x: x[2])

    df['NoCost Rate (30yr)'] = df['New Rate (30yr)'] + NO_COST_RATE_ADJ
    df['Roll Rate (30yr)'] = df['New Rate (30yr)'] + ROLL_RATE_ADJ
    df['CashOut Rate (30yr)'] = df['New Rate (30yr)'] + CASHOUT_RATE_ADJ

    progress_bar.progress(0.5)
    for term, months in [('15yr', 180), ('20yr', 240), ('30yr', 360)]:
        df[f'Pmt Reg ({term})'] = df.apply(lambda row: calculate_pmt(row[f'New Rate ({term})']/12, months, row['New Loan Balance']), axis=1)
        df[f'Pmt NoCost ({term})'] = df.apply(lambda row: calculate_pmt(row['NoCost Rate (30yr)']/12, months, row['New Loan Balance']), axis=1) if term == '30yr' else np.nan
        df[f'Pmt Roll ({term})'] = df.apply(lambda row: calculate_pmt(row['Roll Rate (30yr)']/12, months, row['New Loan Balance']), axis=1) if term == '30yr' else np.nan

    df['Max CashOut Amount'] = (df['New Estimated Home Value'] * MAX_CASHOUT_LTV) - df['New Loan Balance']
    df['Max CashOut Amount'] = df['Max CashOut Amount'].clip(lower=0)
    df['CashOut Loan'] = df['New Loan Balance'] + df['Max CashOut Amount']
    for term, months in [('15yr', 180), ('20yr', 240), ('30yr', 360)]:
        df[f'Pmt CashOut ({term})'] = df.apply(lambda row: calculate_pmt(row['CashOut Rate (30yr)']/12, months, row['CashOut Loan']), axis=1)

    for opt in ['Reg', 'NoCost', 'Roll', 'CashOut']:
        for term in ['15yr', '20yr', '30yr']:
            if f'Pmt {opt} ({term})' in df.columns:
                df[f'Savings {opt} ({term})'] = df['Current P&I Mtg Pymt'] - df[f'Pmt {opt} ({term})']

    df['Equity Increase ($)'] = df['New Estimated Home Value'] - df['Original Appraised Value']

    # Round
    currency_cols = [col for col in df.columns if any(k in col.lower() for k in ['pmt', 'savings', 'amount', 'balance', 'value', 'equity', 'loan'])]
    df[currency_cols] = df[currency_cols].round(2)
    rate_cols = [col for col in df.columns if 'rate' in col.lower() or 'ltv' in col.lower()]
    df[rate_cols] = df[rate_cols].round(5)

    df['Validation'] = df.apply(validate_pmt, axis=1)

    progress_bar.progress(0.7)
    texts_list = batch_generate_texts(df, officer_name, company_name, app_link)

    progress_bar.progress(0.8)
    word_reports = []
    pdf_backups = []
    email_templates = []
    for idx, (row, texts) in enumerate(zip(df.iterrows(), texts_list)):
        _, borrower_data = row
        chart_buf = generate_savings_chart(borrower_data)
        word_path = generate_word_report(borrower_data, texts, officer_name, company_name, chart_buf)
        generate_pdf_backup(borrower_data, texts, officer_name, company_name, chart_buf)
        word_reports.append(word_path)
        email_templates.append(generate_email_template(borrower_data, texts, officer_name, company_name))

    df['Email Template'] = email_templates
    df = df.sort_values(by='Savings Reg (30yr)', ascending=False)  # Prioritize high-savings

    # Excel Output with Formatting
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    action_sheet_path = os.path.join(OUTPUT_FOLDER, f"{officer_name.replace(' ', '_')}_{company_name.replace(' ', '_')}_{timestamp}_ACTION_SHEET.xlsx")
    with pd.ExcelWriter(action_sheet_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Action Sheet', index=False)
        worksheet = writer.sheets['Action Sheet']
        # Conditional formatting for savings (green positive, red negative)
        worksheet.conditional_format(1, df.columns.get_loc('Savings Reg (30yr)'), len(df), df.columns.get_loc('Savings CashOut (30yr)'), {'type': 'cell', 'criteria': '>', 'value': 0, 'format': writer.book.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})})
        worksheet.conditional_format(1, df.columns.get_loc('Savings Reg (30yr)'), len(df), df.columns.get_loc('Savings CashOut (30yr)'), {'type': 'cell', 'criteria': '<', 'value': 0, 'format': writer.book.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})})
        # Summary sheet
        summary_df = pd.DataFrame({
            'Metric': ['Average Savings (Reg 30yr)', 'Total Potential Cash-Out', 'Top Borrower by Savings'],
            'Value': [df['Savings Reg (30yr)'].mean(), df['Max CashOut Amount'].sum(), df.iloc[0]['Borrower First Name']]
        })
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

    # CRM Export
    df.to_json(os.path.join(OUTPUT_FOLDER, 'crm_export.json'), orient='records')

    progress_bar.progress(1.0)
    return df, word_reports

# Streamlit App
st.title("Real Refi Loan Automation Dashboard")
st.write("Upload borrower data to generate personalized refinance options, reports, and outreach texts.")

uploaded_file = st.file_uploader("Upload Borrower Excel", type=['xlsx'])
officer_name = st.text_input("Loan Officer Name", "Michael Young")
company_name = st.text_input("Company Name", "Universal Financial Mortgage")
app_link = st.text_input("Application Link", "https://example.com/apply")

if uploaded_file and st.button("Process Loans"):
    progress_bar = st.progress(0)
    df = pd.read_excel(uploaded_file)
    expected_cols = ['Borrower First Name', 'Borrower Last Name', 'Subject Property Address', 'Subject Property City', 'Subject Property State', 'Total Original Loan Amount', 'Original Appraised Value', 'First Pymt Date', 'Current Interest Rate', 'Loan Term (years)', 'Current P&I Mtg Pymt', 'Borr Cell', 'Borr Email']
    missing_cols = [col for col in expected_cols if col not in df.columns]
    if missing_cols:
        st.error(f"Missing columns: {', '.join(missing_cols)}. Use the template.")
    else:
        df, reports = process_loans(df, officer_name, company_name, app_link, progress_bar)
        st.dataframe(df.style.background_gradient(subset=[col for col in df.columns if 'Savings' in col], cmap='RdYlGn'))
        avg_savings = df['Savings Reg (30yr)'].mean()
        st.metric("Average Monthly Savings (Reg 30yr)", f"${avg_savings:.2f}")
        st.download_button("Download Action Sheet (Excel)", open(os.path.join(OUTPUT_FOLDER, action_sheet_path), 'rb').read(), f"{action_sheet_path.split('/')[-1]}")
        for report in reports:
            st.download_button(f"Download Word Report: {os.path.basename(report)}", open(report, 'rb').read(), os.path.basename(report))
        st.success("Processing complete. All features (reports, texts, simulator) are ready.")

# Simulator
st.subheader("Custom Refinance Simulator")
if 'df' in locals():
    selected = st.selectbox("Select Borrower", df['Borrower First Name'])
    if selected:
        row = df[df['Borrower First Name'] == selected].iloc[0]
        custom_rate = st.slider("Custom Rate (%)", 0.0, 10.0, row['New Rate (30yr)'] * 100) / 100
        custom_term = st.slider("Custom Term (Years)", 10, 30, 30)
        custom_cashout = st.number_input("Custom Cash-Out ($)", 0.0, float(row['Max CashOut Amount']), 0.0)
        custom_pmt, custom_ltv = simulate_scenario(row['New Loan Balance'], row['New Estimated Home Value'], custom_rate, custom_term, custom_cashout)
        custom_savings = row['Current P&I Mtg Pymt'] - custom_pmt
        st.write(f"Custom PMT: ${custom_pmt:.2f} | LTV: {custom_ltv:.2%} | Savings: ${custom_savings:.2f}")
        if st.button("Save Custom Scenario"):
            custom_df = pd.DataFrame([{'Borrower': selected, 'Custom Rate': custom_rate, 'Custom Term': custom_term, 'Custom CashOut': custom_cashout, 'Custom PMT': custom_pmt, 'Custom Savings': custom_savings}])
            custom_path = os.path.join(OUTPUT_FOLDER, f"{selected}_custom_scenario.xlsx")
            custom_df.to_excel(custom_path, index=False)
            st.download_button("Download Custom Scenario", open(custom_path, 'rb').read(), os.path.basename(custom_path))