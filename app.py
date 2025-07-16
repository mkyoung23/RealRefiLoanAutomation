import streamlit as st
import pandas as pd
import google.generativeai as genai
import json
from datetime import datetime
import re
import io

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="MyMCMB AI Command Center",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- STYLING ---
st.markdown("""
<style>
    .stApp {
        background-color: #f0f2f6;
    }
    .main .block-container {
        padding-top: 2rem;
    }
    h1, h2, h3 {
        color: #1e3a8a; /* Dark Blue */
        font-weight: 700;
    }
    .stButton>button {
        background-color: #1e3a8a;
        color: white;
        border-radius: 0.5rem;
        border: none;
        padding: 0.75rem 1.5rem;
        font-weight: 600;
        transition: all 0.2s ease-in-out;
    }
    .stButton>button:hover {
        background-color: #1e40af;
        transform: translateY(-2px);
    }
</style>
""", unsafe_allow_html=True)

# --- API & MODEL SETUP ---
try:
    GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", "")
    if not GEMINI_API_KEY:
        st.error("CRITICAL ERROR: Your Gemini API Key is not configured. Please add it to your Streamlit Secrets.")
        st.stop()
    genai.configure(api_key=GEMINI_API_KEY)
    model = genai.GenerativeModel('gemini-1.5-flash')
except Exception as e:
    st.error(f"Could not configure AI models. Error: {e}")
    st.stop()

# --- SIDEBAR & NAVIGATION ---
st.sidebar.title("MyMCMB AI Command Center")
app_mode = st.sidebar.selectbox(
    "Select an AI Agent",
    ["Refinance Intelligence Center", "Guideline & Product Chatbot", "Social Media Automation", "Admin Rate Panel"]
)

# --- SHARED FUNCTIONS ---
def clean_currency(value):
    if pd.isna(value): return 0.0
    try:
        return float(re.sub(r'[$,]', '', str(value)).strip())
    except (ValueError, TypeError):
        return 0.0

def clean_percentage(value):
    if pd.isna(value): return 0.0
    try:
        val = float(str(value).replace('%', '').strip())
        return val / 100 if val > 1 else val
    except (ValueError, TypeError):
        return 0.0

def calculate_new_pi(principal, annual_rate, term_years):
    principal = clean_currency(principal)
    annual_rate = clean_percentage(annual_rate)
    term_years = int(term_years)
    monthly_rate = annual_rate / 12
    num_payments = term_years * 12
    if monthly_rate <= 0 or num_payments <= 0: return 0.0
    try:
        payment = principal * (monthly_rate * (1 + monthly_rate)**num_payments) / ((1 + monthly_rate)**num_payments - 1)
        return round(payment, 2)
    except (ValueError, ZeroDivisionError):
        return 0.0

@st.cache_data(ttl=3600) # Cache for 1 hour
def get_estimated_property_value(address):
    # This is a placeholder for a real API call. In a real scenario, you'd use a paid service.
    # For this demo, we'll simulate a lookup with a calculation.
    # A real implementation would look like:
    # response = requests.get(f"https://api.realestate.com?address={address}&key=API_KEY")
    # return response.json()['value']
    # For now, we'll just return a calculated value. This part can be upgraded later.
    return None # Indicates that we should use the appreciation method

def calculate_amortized_balance(principal, annual_rate, term_years, first_payment_date):
    principal = clean_currency(principal)
    annual_rate = clean_percentage(annual_rate)
    term_years = int(term_years)
    if pd.isna(first_payment_date): return principal
    try:
        first_payment = pd.to_datetime(first_payment_date)
        months_elapsed = (datetime.now().year - first_payment.year) * 12 + (datetime.now().month - first_payment.month)
        payments_made = max(0, months_elapsed)
        if payments_made == 0: return principal
        monthly_rate = annual_rate / 12
        total_payments = term_years * 12
        if monthly_rate <= 0: return principal * (1 - (payments_made / total_payments))
        balance = principal * ( ((1 + monthly_rate)**total_payments - (1 + monthly_rate)**payments_made) / ((1 + monthly_rate)**total_payments - 1) )
        return max(0, round(balance, 2))
    except Exception: return principal

# --- ADMIN RATE PANEL ---
if app_mode == "Admin Rate Panel":
    st.title("Admin Rate Panel")
    st.write("Set the current mortgage rates that the Refinance agent will use for calculations.")

    if 'rates' not in st.session_state:
        st.session_state.rates = {
            '30yr_fixed': 6.875, '20yr_fixed': 6.625, '15yr_fixed': 6.000,
            '5yr_arm': 7.394, 'no_cost_adj': 0.250
        }

    with st.form("rate_form"):
        st.subheader("Current Market Rates (%)")
        rates = st.session_state.rates
        rates['30yr_fixed'] = st.number_input("30-Year Fixed Rate", value=rates['30yr_fixed'], format="%.3f")
        rates['20yr_fixed'] = st.number_input("20-Year Fixed Rate", value=rates['20yr_fixed'], format="%.3f")
        rates['15yr_fixed'] = st.number_input("15-Year Fixed Rate", value=rates['15yr_fixed'], format="%.3f")
        rates['5yr_arm'] = st.number_input("5/1 ARM Rate", value=rates['5yr_arm'], format="%.3f")
        rates['no_cost_adj'] = st.number_input("No-Cost Rate Adjustment", value=rates['no_cost_adj'], format="%.3f", help="Amount to add to the 30yr rate for a no-cost option.")
        
        submitted = st.form_submit_button("Save Rates")
        if submitted:
            st.session_state.rates = rates
            st.success("Rates updated successfully!")

# --- REFINANCE INTELLIGENCE CENTER ---
elif app_mode == "Refinance Intelligence Center":
    st.title("Refinance Intelligence Center")
    st.markdown("### Upload a borrower data sheet to generate hyper-personalized outreach plans.")

    uploaded_file = st.file_uploader("Choose a borrower Excel file", type=['xlsx'])

    if uploaded_file:
        df_original = pd.read_excel(uploaded_file, engine='openpyxl')
        st.success(f"Successfully loaded {len(df_original)} borrowers from '{uploaded_file.name}'.")

        if st.button("🚀 Generate AI Outreach Plans"):
            with st.spinner("Initiating AI Analysis... This will take a few moments."):
                df = df_original.copy()
                rates = st.session_state.get('rates', {'30yr_fixed': 6.875, '20yr_fixed': 6.625, '15yr_fixed': 6.000, '5yr_arm': 7.394, 'no_cost_adj': 0.250})

                # --- Data Processing Pipeline ---
                progress_bar = st.progress(0, text="Calculating financial scenarios...")
                df['Remaining Balance'] = df.apply(lambda row: calculate_amortized_balance(row.get('Total Original Loan Amount'), row.get('Current Interest Rate'), row.get('Loan Term (years)'), row.get('First Pymt Date')), axis=1)
                df['Months Since First Payment'] = df['First Pymt Date'].apply(lambda x: max(0, (datetime.now().year - pd.to_datetime(x).year) * 12 + (datetime.now().month - pd.to_datetime(x).month)) if pd.notna(x) else 0)
                df['Estimated Home Value'] = df.apply(lambda row: round(clean_currency(row.get('Original Property Value', 0)) * (1.04 ** (row['Months Since First Payment'] / 12)), 2), axis=1)
                df['Estimated LTV'] = (df['Remaining Balance'] / df['Estimated Home Value']).fillna(0)
                df['Max Cash-Out Amount'] = (df['Estimated Home Value'] * 0.80) - df['Remaining Balance']
                df['Max Cash-Out Amount'] = df['Max Cash-Out Amount'].apply(lambda x: max(0, round(x, 2)))
                
                # Scenario Calculations
                for term, rate_key in [('30yr', '30yr_fixed'), ('20yr', '20yr_fixed'), ('15yr', '15yr_fixed')]:
                    rate = rates[rate_key] / 100
                    df[f'New P&I ({term})'] = df['Remaining Balance'].apply(lambda bal: calculate_new_pi(bal, rate, int(term.replace('yr',''))))
                    df[f'Savings ({term})'] = df['Current P&I Mtg Pymt'].apply(clean_currency) - df[f'New P&I ({term})']
                
                # --- AI Content Generation ---
                outreach_results = []
                for i, row in df.iterrows():
                    progress_bar.progress((i + 1) / len(df), text=f"Generating AI outreach for {row['Borrower First Name']}...")
                    
                    # Same Payment Cash-Out Calculation
                    current_payment = clean_currency(row['Current P&I Mtg Pymt'])
                    new_rate = rates['30yr_fixed'] / 100
                    # Simplified calculation for new loan amount with same payment
                    try:
                        max_loan_for_same_payment = (current_payment * (((1 + new_rate/12)**360) - 1)) / ((new_rate/12) * (1 + new_rate/12)**360)
                        cash_out_same_payment = max(0, max_loan_for_same_payment - row['Remaining Balance'])
                    except (ZeroDivisionError, ValueError):
                        cash_out_same_payment = 0

                    prompt = f"""
                    You are an expert mortgage loan officer assistant for MyMCMB. Your task is to generate a set of personalized, human-sounding outreach messages for a past client named {row['Borrower First Name']}. The tone must be professional, helpful, and sound like it came from a real person.

                    **Borrower's Financial Snapshot:**
                    - Property City: {row.get('City', 'their city')}
                    - Current Monthly P&I: ${clean_currency(row['Current P&I Mtg Pymt']):.2f}
                    - Estimated Home Value: ${row['Estimated Home Value']:.2f}
                    
                    **Calculated Refinance Scenarios:**
                    1.  **30-Year Fixed:** New Payment: ${row['New P&I (30yr)']:.2f}, Monthly Savings: ${row['Savings (30yr)']:.2f}
                    2.  **15-Year Fixed:** New Payment: ${row['New P&I (15yr)']:.2f}, Monthly Savings: ${row['Savings (15yr)']:.2f}
                    3.  **Max Cash-Out:** You can offer up to ${row['Max Cash-Out Amount']:.2f} in cash.
                    4.  **"Same Payment" Cash-Out:** You can offer approx. ${cash_out_same_payment:.2f} in cash while keeping their payment nearly the same.

                    **Task:**
                    Generate a JSON object with four distinct outreach options. Each option should have a 'title', a concise 'sms' template, and a professional 'email' template.
                    1.  **"Significant Savings Alert"**: Focus on the 30-year option's direct monthly savings.
                    2.  **"Aggressive Payoff Plan"**: Focus on the 15-year option, highlighting owning their home faster.
                    3.  **"Leverage Your Equity"**: Focus on the maximum cash-out option for home improvements or debt consolidation.
                    4.  **"Cash with No Payment Shock"**: Focus on the 'same payment' cash-out option.
                    Make the messages sound authentic. For one of the emails, mention a positive local event or trend in {row.get('City', 'their area')} to personalize it further.
                    """
                    try:
                        response = model.generate_content(prompt, generation_config=genai.types.GenerationConfig(response_mime_type="application/json"))
                        outreach_results.append(json.loads(response.text))
                    except Exception:
                        outreach_results.append({"outreach_options": []})

                df['AI_Outreach'] = outreach_results
                st.session_state.df_results = df
                st.success("Analysis complete! View the outreach plans below.")

        if 'df_results' in st.session_state:
            st.markdown("---")
            st.header("Generated Outreach Blueprints")
            df_results = st.session_state.df_results
            
            # Display results in the app
            for index, row in df_results.iterrows():
                with st.expander(f"👤 **{row['Borrower First Name']} {row.get('Borrower Last Name', '')}** | Max Savings: **${row['Savings (30yr)']:.2f}/mo**"):
                    st.subheader("Financial Snapshot")
                    # Display metrics...
                    if row['AI_Outreach'] and row['AI_Outreach'].get('outreach_options'):
                        for option in row['AI_Outreach']['outreach_options']:
                            st.markdown(f"#### {option.get('title', 'Outreach Option')}")
                            # Display SMS/Email text areas...
                    else:
                        st.warning("Could not generate outreach content.")

# --- OTHER AGENTS ---
elif app_mode == "Guideline & Product Chatbot":
    st.title("Guideline & Product Chatbot")
    st.info("This AI Agent is under construction.")

elif app_mode == "Social Media Automation":
    st.title("Social Media Automation")
    st.info("This AI Agent is under construction.")
