import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import datetime
import calendar
import holidays
import re
import os
import altair as alt

# -------------------------------------------------------------
# PAGE CONFIGURATION & THEME
# -------------------------------------------------------------
st.set_page_config(
    page_title="Fathima Traders — Payroll & Settlement Portal",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------------------------------------------
# HTML RENDERING HELPER (PREVENTS MARKDOWN INDENTATION BUGS)
# -------------------------------------------------------------
def render_clean_html(html_str):
    """Strip all leading spaces on every line so Markdown never treats HTML as code blocks."""
    lines = [line.lstrip() for line in html_str.strip().splitlines()]
    st.markdown("\n".join(lines), unsafe_allow_html=True)

# Global styles injection
render_clean_html("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }

    /* Hero Header */
    .hero-box {
        background: linear-gradient(135deg, #0B132B 0%, #1C2541 50%, #1E3A8A 100%);
        border-radius: 16px;
        padding: 1.8rem 2.2rem;
        color: white;
        margin-bottom: 1.5rem;
        box-shadow: 0 10px 25px -5px rgba(11, 19, 43, 0.3);
        border: 1px solid rgba(255, 255, 255, 0.1);
    }
    .hero-title {
        font-size: 2.1rem;
        font-weight: 800;
        letter-spacing: -0.02em;
        margin: 0;
        color: #FFFFFF;
    }
    .hero-subtitle {
        font-size: 0.95rem;
        color: #94A3B8;
        margin-top: 0.3rem;
        margin-bottom: 1.1rem;
    }
    .chip-container {
        display: flex;
        flex-wrap: wrap;
        gap: 0.6rem;
    }
    .hero-chip {
        display: inline-flex;
        align-items: center;
        background: rgba(255, 255, 255, 0.09);
        backdrop-filter: blur(8px);
        padding: 0.35rem 0.8rem;
        border-radius: 9999px;
        font-size: 0.8rem;
        color: #E2E8F0;
        border: 1px solid rgba(255, 255, 255, 0.12);
        font-weight: 500;
    }
    .hero-chip strong {
        color: #FFFFFF;
        margin-right: 0.25rem;
    }
    .chip-verified {
        background: rgba(16, 185, 129, 0.2);
        color: #34D399;
        border: 1px solid rgba(52, 211, 153, 0.3);
        font-weight: 600;
    }

    /* Sub-bar stats */
    .summary-bar {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 0.85rem 1.4rem;
        display: flex;
        flex-wrap: wrap;
        justify-content: space-between;
        gap: 1rem;
        margin-bottom: 1.5rem;
    }
    .stat-node {
        text-align: center;
    }
    .stat-node-val {
        font-size: 1.25rem;
        font-weight: 800;
        color: #0F172A;
    }
    .stat-node-lbl {
        font-size: 0.72rem;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        font-weight: 700;
        margin-top: 0.15rem;
    }
</style>
""")

# -------------------------------------------------------------
# HELPER FUNCTIONS
# -------------------------------------------------------------
def extract_employee_name_from_sales(file_obj):
    """Read first 7 rows to extract user name from Sales Register."""
    try:
        df_head = pd.read_excel(file_obj, header=None, nrows=7)
        for val in df_head.values.flatten():
            if pd.notna(val) and re.search(r'User\s*:', str(val), re.IGNORECASE):
                m = re.search(r'User\s*:\s*([^:\n\r]+?)(?=\s+Branch\s*:|\s+From|\s+To|$)', str(val), re.IGNORECASE)
                if m:
                    return m.group(1).strip()
    except Exception as e:
        st.error(f"Error reading Sales Register header: {e}")
    return None

def extract_employee_name_from_ledger(file_obj):
    """Read first 7 rows to extract account name from Ledger."""
    try:
        df_head = pd.read_excel(file_obj, header=None, nrows=7)
        for val in df_head.values.flatten():
            if pd.notna(val) and re.search(r'Account\s*:', str(val), re.IGNORECASE):
                m = re.search(r'Account\s*:\s*([^:\n\r]+?)(?=\s+Branch\s*:|\s+From|\s+To|$)', str(val), re.IGNORECASE)
                if m:
                    return m.group(1).strip()
    except Exception as e:
        st.error(f"Error reading Ledger header: {e}")
    return None

def normalize_name(name):
    """Ignore prefixes like 'FT ' and trim/uppercase for comparison."""
    if not name:
        return ""
    cleaned = re.sub(r'^FT\s+', '', name.strip(), flags=re.IGNORECASE)
    return cleaned.strip().upper()

def load_sales_data(file_obj):
    """Load Sales Register data, skipping the first 7 rows (headers start row 8)."""
    try:
        df = pd.read_excel(file_obj, skiprows=8)
        if 'Date' not in df.columns:
            if hasattr(file_obj, 'seek'):
                file_obj.seek(0)
            df = pd.read_excel(file_obj, skiprows=7)
            if 'Date' not in df.columns and len(df) > 0 and 'Date' in df.iloc[0].values:
                df.columns = df.iloc[0]
                df = df.iloc[1:].reset_index(drop=True)
    except Exception as e:
        st.error(f"Failed to load Sales Register data: {e}")
        return None

    if 'Date' in df.columns:
        df = df[df['Date'].notna()].copy()
        df['Date_parsed'] = pd.to_datetime(df['Date'], errors='coerce')
        df = df[df['Date_parsed'].notna()].copy()
        df['Date_only'] = df['Date_parsed'].dt.date
        if 'Created Time' in df.columns:
            df['Created_Time_parsed'] = pd.to_datetime(df['Created Time'], errors='coerce')
        if 'Bill Amount' in df.columns:
            df['Bill Amount'] = pd.to_numeric(df['Bill Amount'], errors='coerce').fillna(0.0)
    return df

def load_ledger_data(file_obj):
    """Load Ledger data, skipping the first 7 rows (headers start row 8)."""
    try:
        df = pd.read_excel(file_obj, skiprows=8)
        if 'Date' not in df.columns and 'Type' not in df.columns:
            if hasattr(file_obj, 'seek'):
                file_obj.seek(0)
            df = pd.read_excel(file_obj, skiprows=7)
            if ('Date' not in df.columns or 'Type' not in df.columns) and len(df) > 0 and 'Type' in df.iloc[0].values:
                df.columns = df.iloc[0]
                df = df.iloc[1:].reset_index(drop=True)
    except Exception as e:
        st.error(f"Failed to load Ledger data: {e}")
        return None

    if 'Date' in df.columns:
        df = df[df['Date'].notna()].copy()
        df['Date_parsed'] = pd.to_datetime(df['Date'], errors='coerce')
        df = df[df['Date_parsed'].notna()].copy()
        df['Date_only'] = df['Date_parsed'].dt.date
    if 'Debit' in df.columns:
        df['Debit_num'] = pd.to_numeric(
            df['Debit'].astype(str).str.replace(',', '').str.strip(), errors='coerce'
        ).fillna(0.0)
    else:
        df['Debit_num'] = 0.0
    return df

def amount_to_inr_words(amount):
    """Convert numerical amount to clean Indian Rupees words format."""
    try:
        val = int(round(float(amount)))
        if val == 0:
            return "Zero Rupees Only"
        if val < 0:
            return f"Negative INR {amount_to_inr_words(abs(val))}"
        units = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten',
                 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen']
        tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety']
        
        def helper(n):
            if n == 0:
                return ''
            elif n < 20:
                return units[n]
            elif n < 100:
                return tens[n // 10] + ('' if n % 10 == 0 else ' ' + units[n % 10])
            elif n < 1000:
                return units[n // 100] + ' Hundred' + ('' if n % 100 == 0 else ' and ' + helper(n % 100))
            elif n < 100000:
                return helper(n // 1000) + ' Thousand' + ('' if n % 1000 == 0 else ' ' + helper(n % 1000))
            elif n < 10000000:
                return helper(n // 100000) + ' Lakh' + ('' if n % 100000 == 0 else ' ' + helper(n % 100000))
            else:
                return helper(n // 10000000) + ' Crore' + ('' if n % 10000000 == 0 else ' ' + helper(n % 10000000))
        
        return "Rupees " + helper(val) + " Only"
    except Exception:
        return f"Rupees {amount:,.2f}"

# -------------------------------------------------------------
# SIDEBAR CONTROLS & CONFIGURATION
# -------------------------------------------------------------
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/money-transfer.png", width=64)
    st.markdown("### **Payroll Command Center**")
    st.caption("Fathima Traders • Chelari, Kerala")
    st.markdown("---")

    st.markdown("#### 💼 **Salary Configuration**")
    base_salary = st.number_input(
        "Base Monthly Salary (₹)",
        min_value=0.0,
        value=20000.0,
        step=500.0,
        format="%.2f",
        help="Monthly gross compensation standard for complete working attendance."
    )

    st.markdown("---")
    st.markdown("#### 📁 **Source Data Files**")
    uploaded_sales = st.file_uploader(
        "Daily Sales Register (.xlsx)",
        type=["xlsx", "xls"],
        help="Sales register with punch dates, times, and bill amounts."
    )
    uploaded_ledger = st.file_uploader(
        "Ledger Of Account (.xlsx)",
        type=["xlsx", "xls"],
        help="Account ledger with cash advances, payments, and staff billing."
    )

    st.markdown("---")
    st.markdown("#### ⚙️ **Audit Filters**")
    filter_active_month_ledger = st.checkbox(
        "Scope Ledger to Active Month",
        value=True,
        help="Filters ledger deductions to strictly match the month and year of the Sales Register."
    )

    st.markdown("<br/>", unsafe_allow_html=True)
    st.info("💡 **Automatic Detection:** If files are not uploaded, the portal automatically loads local files found in the directory.")

# -------------------------------------------------------------
# FILE RESOLUTION & VALIDATION
# -------------------------------------------------------------
sales_file = uploaded_sales
ledger_file = uploaded_ledger

if sales_file is None and os.path.exists("Daily Sales Register.xlsx"):
    sales_file = "Daily Sales Register.xlsx"

if ledger_file is None:
    if os.path.exists("Ledger Of Account.xlsx"):
        ledger_file = "Ledger Of Account.xlsx"
    elif os.path.exists("Ledger Of Account - FT FARIS.xlsx"):
        ledger_file = "Ledger Of Account - FT FARIS.xlsx"

if sales_file is None or ledger_file is None:
    st.info("👋 Please upload both **Daily Sales Register.xlsx** and **Ledger Of Account.xlsx** via the sidebar.")
    st.stop()

# -------------------------------------------------------------
# 1. EMPLOYEE NAME EXTRACTION & VALIDATION (CRITICAL)
# -------------------------------------------------------------
name_a = extract_employee_name_from_sales(sales_file)
name_b = extract_employee_name_from_ledger(ledger_file)

if not name_a or not name_b:
    st.error(f"Failed to extract employee identity from headers: Sales Register User='{name_a}', Ledger Account='{name_b}'.")
    st.stop()

norm_a = normalize_name(name_a)
norm_b = normalize_name(name_b)

if norm_a != norm_b:
    st.error(f"Employee mismatch: Sales Register belongs to [{name_a}], but Ledger belongs to [{name_b}].")
    st.stop()

clean_emp_display = name_a if name_a.upper() == norm_a else norm_a

# -------------------------------------------------------------
# 2. DATA PROCESSING & DYNAMIC DATES
# -------------------------------------------------------------
df_sales = load_sales_data(sales_file)
df_ledger = load_ledger_data(ledger_file)

if df_sales is None or df_sales.empty:
    st.error("Sales Register file contains no valid sales records with dates.")
    st.stop()

if df_ledger is None or df_ledger.empty:
    st.error("Ledger file contains no valid financial transactions.")
    st.stop()

# Active Month & Year Detection
active_year = int(df_sales['Date_parsed'].dt.year.mode()[0])
active_month = int(df_sales['Date_parsed'].dt.month.mode()[0])
month_name = calendar.month_name[active_month]
total_days_in_month = calendar.monthrange(active_year, active_month)[1]
max_date_in_file = df_sales['Date_only'].max()

# -------------------------------------------------------------
# 3. HOLIDAYS & OFF-DAYS
# -------------------------------------------------------------
kl_holidays = holidays.country_holidays('IN', subdiv='KL', years=active_year)

paid_holiday_dates = set()
holiday_names = {}

for day in range(1, total_days_in_month + 1):
    d = datetime.date(active_year, active_month, day)
    is_sunday = (d.weekday() == 6)
    is_pub_holiday = (d in kl_holidays)
    if is_sunday or is_pub_holiday:
        paid_holiday_dates.add(d)
        if is_pub_holiday:
            holiday_names[d] = kl_holidays.get(d)
        else:
            holiday_names[d] = "Sunday"

total_paid_holidays = len(paid_holiday_dates)

# -------------------------------------------------------------
# 4. ATTENDANCE & TIME RULES (DYNAMIC MONTH ITERATION)
# -------------------------------------------------------------
attendance_records = []
total_unpaid_equivalent = 0.0

for day in range(1, total_days_in_month + 1):
    current_date = datetime.date(active_year, active_month, day)
    day_sales = df_sales[df_sales['Date_only'] == current_date]
    has_activity = len(day_sales) > 0

    first_time_str = "-"
    last_time_str = "-"
    last_time_val = None

    if has_activity and 'Created_Time_parsed' in day_sales.columns:
        valid_times = day_sales['Created_Time_parsed'].dropna()
        if len(valid_times) > 0:
            first_time_str = valid_times.min().strftime('%I:%M %p')
            last_time_val = valid_times.max()
            last_time_str = last_time_val.strftime('%I:%M %p')

    # Condition A: Paid Holiday
    if current_date in paid_holiday_dates:
        h_label = holiday_names.get(current_date, "Holiday")
        status = f"Paid Holiday ({h_label})"
        badge_bg = "#DBEAFE"
        badge_color = "#1E40AF"
        unpaid = 0.0
    # Condition B: Worked Day
    elif has_activity:
        cutoff_time = datetime.time(14, 0, 0)
        if last_time_val is not None and last_time_val.time() <= cutoff_time:
            status = "Half Day"
            badge_bg = "#FEF3C7"
            badge_color = "#92400E"
            unpaid = 0.5
        else:
            status = "Full Day / Present"
            badge_bg = "#DCFCE7"
            badge_color = "#166534"
            unpaid = 0.0
    # Condition C: Incomplete Future Data
    elif current_date > max_date_in_file:
        status = "Present (Assumed)"
        badge_bg = "#F3E8FF"
        badge_color = "#6B21A8"
        unpaid = 0.0
    # Condition D: Absent
    else:
        status = "Leave / Absent"
        badge_bg = "#FEE2E2"
        badge_color = "#991B1B"
        unpaid = 1.0

    total_unpaid_equivalent += unpaid
    bills_count = len(day_sales)
    total_sales_val = day_sales['Bill Amount'].sum() if has_activity else 0.0

    attendance_records.append({
        'Date': current_date,
        'Day_Num': day,
        'Day_Name': current_date.strftime('%a'),
        'Status': status,
        'Badge_BG': badge_bg,
        'Badge_Color': badge_color,
        'Unpaid Eq': unpaid,
        'Earliest Punch': first_time_str,
        'Latest Punch': last_time_str,
        'Bills Count': bills_count,
        'Total Sales (₹)': total_sales_val
    })

df_attendance = pd.DataFrame(attendance_records)

# -------------------------------------------------------------
# 5. LEDGER DEDUCTIONS
# -------------------------------------------------------------
df_ledger_scoped = df_ledger.copy()
if filter_active_month_ledger:
    df_ledger_scoped = df_ledger_scoped[
        (df_ledger_scoped['Date_parsed'].dt.month == active_month) &
        (df_ledger_scoped['Date_parsed'].dt.year == active_year)
    ].copy()

deductions = []
ignored_entries = []

for _, row in df_ledger_scoped.iterrows():
    l_type = str(row.get('Type', '')).strip()
    l_remarks = str(row.get('Remarks', '')).strip()
    debit_amt = float(row.get('Debit_num', 0.0))

    if "SALARY CASH" in l_remarks.upper() or "OPENING BALANCE" in l_remarks.upper() or "OPENING BALANCE" in l_type.upper():
        ignored_entries.append({
            'Doc No': row.get('Doc No', '-'),
            'Date': row['Date_only'],
            'Type': l_type,
            'Remarks': l_remarks,
            'Amount': debit_amt,
            'Reason': "Excluded per policy (Salary disbursement / Opening Balance)"
        })
        continue

    if l_type.upper() == "PAYMENT" and "ADVANCE CASH" in l_remarks.upper():
        deductions.append({
            'Doc No': row.get('Doc No', '-'),
            'Date': row['Date_only'],
            'Type': l_type,
            'Category': 'Advance',
            'Remarks': l_remarks,
            'Amount (₹)': debit_amt
        })
    elif l_type.upper() == "SALES":
        deductions.append({
            'Doc No': row.get('Doc No', '-'),
            'Date': row['Date_only'],
            'Type': l_type,
            'Category': 'Staff Sales',
            'Remarks': l_remarks,
            'Amount (₹)': debit_amt
        })

df_deductions = pd.DataFrame(deductions)
if df_deductions.empty:
    df_deductions = pd.DataFrame(columns=['Doc No', 'Date', 'Type', 'Category', 'Remarks', 'Amount (₹)'])

total_advances = df_deductions[df_deductions['Category'] == 'Advance']['Amount (₹)'].sum() if not df_deductions.empty else 0.0
total_staff_sales = df_deductions[df_deductions['Category'] == 'Staff Sales']['Amount (₹)'].sum() if not df_deductions.empty else 0.0
total_ledger_deductions = total_advances + total_staff_sales

# -------------------------------------------------------------
# 6. PAYROLL MATH CALCULATION
# -------------------------------------------------------------
working_days = max(1, total_days_in_month - total_paid_holidays)
daily_wage = base_salary / working_days
attendance_deduction = total_unpaid_equivalent * daily_wage
net_payable = base_salary - attendance_deduction - total_advances - total_staff_sales
net_paid_days = working_days - total_unpaid_equivalent
payout_pct = (net_payable / base_salary * 100) if base_salary > 0 else 0

# -------------------------------------------------------------
# TOP HERO BANNER
# -------------------------------------------------------------
hero_html = f"""
<div class="hero-box">
    <div style="display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 1rem;">
        <div>
            <h1 class="hero-title">Fathima Traders — Payroll & Settlement</h1>
            <div class="hero-subtitle">Comprehensive Attendance Audit, Store Revenue Attribution & Wage Settlement Engine</div>
            <div class="chip-container">
                <span class="hero-chip"><strong>Staff:</strong> {clean_emp_display}</span>
                <span class="hero-chip"><strong>Ledger:</strong> {name_b}</span>
                <span class="hero-chip"><strong>Period:</strong> {month_name} {active_year}</span>
                <span class="hero-chip"><strong>Records Thru:</strong> {max_date_in_file.strftime('%d %b %Y')}</span>
                <span class="hero-chip chip-verified">✓ Identity Cross-Verified</span>
            </div>
        </div>
        <div style="text-align: right; background: rgba(255,255,255,0.08); padding: 0.8rem 1.4rem; border-radius: 12px; border: 1px solid rgba(255,255,255,0.15);">
            <div style="font-size: 0.72rem; text-transform: uppercase; color: #94A3B8; letter-spacing: 0.08em; font-weight: 700;">Final Net Settlement</div>
            <div style="font-size: 2.3rem; font-weight: 800; color: #4ADE80; line-height: 1.1;">₹{net_payable:,.2f}</div>
            <div style="font-size: 0.76rem; color: #E2E8F0; margin-top: 0.2rem;">Net Payable Payout ({payout_pct:.1f}%)</div>
        </div>
    </div>
</div>
"""
render_clean_html(hero_html)

# -------------------------------------------------------------
# 7. EXECUTIVE KPI CARDS
# -------------------------------------------------------------
k1, k2, k3, k4 = st.columns(4)

with k1:
    st.metric(
        label="Gross Base Salary",
        value=f"₹{base_salary:,.2f}",
        delta=f"₹{daily_wage:,.2f} / Day Rate",
        help=f"Allocated over {working_days} working days"
    )

with k2:
    st.metric(
        label="Attendance LOP Deduction",
        value=f"- ₹{attendance_deduction:,.2f}",
        delta=f"{total_unpaid_equivalent:.1f} Unpaid Days",
        delta_color="inverse",
        help=f"{total_unpaid_equivalent} days × ₹{daily_wage:,.2f}/day"
    )

with k3:
    st.metric(
        label="Ledger Deductions",
        value=f"- ₹{total_ledger_deductions:,.2f}",
        delta=f"Adv: ₹{total_advances:,.0f} | Sales: ₹{total_staff_sales:,.0f}",
        delta_color="inverse",
        help="Cash advances + Staff personal billing debits"
    )

with k4:
    net_color = "normal" if net_payable >= 0 else "inverse"
    st.metric(
        label="Final NET PAYABLE",
        value=f"₹{net_payable:,.2f}",
        delta=f"{net_paid_days:.1f} Effective Paid Days",
        delta_color=net_color,
        help="Final amount due to employee after all deductions"
    )

# Working Days Quick Bar
summary_bar_html = f"""
<div class="summary-bar">
    <div class="stat-node">
        <div class="stat-node-val">{total_days_in_month}</div>
        <div class="stat-node-lbl">Month Days</div>
    </div>
    <div class="stat-node">
        <div class="stat-node-val" style="color: #2563EB;">{total_paid_holidays}</div>
        <div class="stat-node-lbl">Paid Holidays</div>
    </div>
    <div class="stat-node">
        <div class="stat-node-val" style="color: #059669;">{working_days}</div>
        <div class="stat-node-lbl">Net Working Days</div>
    </div>
    <div class="stat-node">
        <div class="stat-node-val">₹{daily_wage:,.2f}</div>
        <div class="stat-node-lbl">Daily Wage Rate</div>
    </div>
    <div class="stat-node">
        <div class="stat-node-val" style="color: #DC2626;">{total_unpaid_equivalent:.1f}</div>
        <div class="stat-node-lbl">Unpaid Absence Days</div>
    </div>
    <div class="stat-node">
        <div class="stat-node-val" style="color: #16A34A;">{net_paid_days:.1f}</div>
        <div class="stat-node-lbl">Net Paid Days</div>
    </div>
</div>
"""
render_clean_html(summary_bar_html)

# -------------------------------------------------------------
# TABBED NAVIGATION
# -------------------------------------------------------------
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🧾 Official Settlement Payslip",
    "🗓️ Attendance & Punch Audit",
    "💳 Ledger & Advances Audit",
    "📈 Sales & Store Performance",
    "🔍 Business Rules & Validation"
])

# -------------------------------------------------------------
# TAB 1: FORMAL PRINTABLE SALARY SETTLEMENT VOUCHER
# -------------------------------------------------------------
with tab1:
    st.markdown("### 📄 **Staff Salary Settlement Voucher**")
    st.caption("Official company payroll statement with earnings, deductions, amount in words, and verification signatures.")

    inr_words = amount_to_inr_words(net_payable)
    voucher_date = datetime.date.today().strftime('%d-%b-%Y')

    # Construct complete standalone HTML for voucher (rendered via components.html)
    voucher_full_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;700&display=swap" rel="stylesheet">
    <style>
        body {{
            font-family: 'Plus Jakarta Sans', system-ui, sans-serif;
            background: #F8FAFC;
            margin: 0;
            padding: 15px;
            color: #0F172A;
        }}
        .voucher-card {{
            background: #FFFFFF;
            border: 2px solid #E2E8F0;
            border-radius: 14px;
            padding: 28px 32px;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.08);
            max-width: 900px;
            margin: 0 auto;
        }}
        .voucher-header {{
            border-bottom: 2px solid #0F172A;
            padding-bottom: 16px;
            margin-bottom: 20px;
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
        }}
        .voucher-company {{
            font-size: 26px;
            font-weight: 800;
            color: #0F172A;
            margin: 0;
            letter-spacing: -0.01em;
        }}
        .voucher-address {{
            font-size: 13px;
            color: #475569;
            margin-top: 4px;
        }}
        .voucher-meta-box {{
            background: #F8FAFC;
            border: 1px solid #CBD5E1;
            border-radius: 8px;
            padding: 8px 14px;
            font-size: 13px;
            text-align: right;
            line-height: 1.5;
        }}
        .info-grid {{
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 12px;
            background: #F8FAFC;
            border: 1px solid #E2E8F0;
            border-radius: 10px;
            padding: 14px 18px;
            margin-bottom: 22px;
            font-size: 13px;
        }}
        .info-label {{
            color: #64748B;
            font-size: 11px;
            text-transform: uppercase;
            font-weight: 700;
            letter-spacing: 0.05em;
        }}
        .info-value {{
            font-weight: 700;
            color: #0F172A;
            font-size: 14px;
            margin-top: 2px;
        }}
        .reconcile-grid {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
        }}
        th {{
            background: #F1F5F9;
            padding: 9px 12px;
            font-size: 12px;
            font-weight: 700;
            color: #1E293B;
            border-bottom: 2px solid #CBD5E1;
            text-align: left;
        }}
        td {{
            padding: 9px 12px;
            font-size: 13px;
            color: #334155;
            border-bottom: 1px solid #E2E8F0;
        }}
        .total-row td {{
            background: #F8FAFC;
            font-weight: 700;
            font-size: 14px;
            border-top: 1.5px solid #CBD5E1;
            border-bottom: 2px solid #CBD5E1;
        }}
        .net-banner {{
            background: #F0FDF4;
            border: 2px solid #86EFAC;
            border-radius: 10px;
            padding: 16px 20px;
            margin-top: 22px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        .net-lbl {{
            font-size: 12px;
            text-transform: uppercase;
            color: #166534;
            font-weight: 700;
            letter-spacing: 0.05em;
        }}
        .net-words {{
            font-size: 14px;
            color: #14532D;
            margin-top: 4px;
            font-weight: 600;
        }}
        .net-amt {{
            font-size: 32px;
            font-weight: 900;
            color: #15803D;
            font-family: 'JetBrains Mono', monospace;
        }}
        .sig-row {{
            display: flex;
            justify-content: space-between;
            margin-top: 50px;
            padding-top: 10px;
        }}
        .sig-box {{
            text-align: center;
            width: 220px;
            border-top: 1.5px dashed #94A3B8;
            padding-top: 6px;
            font-size: 12px;
            color: #64748B;
            font-weight: 600;
        }}
        .btn-print {{
            background: #0F172A;
            color: white;
            border: none;
            padding: 8px 18px;
            border-radius: 6px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            margin-bottom: 15px;
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }}
        .btn-print:hover {{
            background: #1E293B;
        }}
        @media print {{
            body {{ background: white; padding: 0; }}
            .btn-print {{ display: none; }}
            .voucher-card {{ border: none; box-shadow: none; padding: 0; }}
        }}
    </style>
    </head>
    <body>
        <div style="max-width: 900px; margin: 0 auto; text-align: right;">
            <button class="btn-print" onclick="window.print()">🖨️ Print / Save as PDF</button>
        </div>
        <div class="voucher-card">
            <div class="voucher-header">
                <div>
                    <h2 class="voucher-company">FATHIMA TRADERS</h2>
                    <div class="voucher-address">CHELARI, MALAPPURAM, KERALA | Phone: 7356153550, 96050005373</div>
                    <div style="font-weight: 800; color: #1E3A8A; margin-top: 6px; font-size: 16px; letter-spacing: 0.02em;">
                        STAFF PAYROLL SETTLEMENT VOUCHER
                    </div>
                </div>
                <div class="voucher-meta-box">
                    <div><strong>Voucher Ref:</strong> FT/PAY/{active_year}/{active_month:02d}</div>
                    <div><strong>Date:</strong> {voucher_date}</div>
                    <div><strong>Period:</strong> {month_name} {active_year}</div>
                </div>
            </div>

            <div class="info-grid">
                <div>
                    <div class="info-label">Employee Name</div>
                    <div class="info-value">{clean_emp_display}</div>
                </div>
                <div>
                    <div class="info-label">Ledger Account</div>
                    <div class="info-value">{name_b}</div>
                </div>
                <div>
                    <div class="info-label">Working Days</div>
                    <div class="info-value">{working_days} Days (excl. {total_paid_holidays} hol.)</div>
                </div>
                <div>
                    <div class="info-label">Daily Wage Rate</div>
                    <div class="info-value">₹{daily_wage:,.2f} / day</div>
                </div>
            </div>

            <div class="reconcile-grid">
                <!-- Earnings Column -->
                <div>
                    <table>
                        <thead>
                            <tr>
                                <th>EARNINGS HEAD</th>
                                <th style="text-align: right;">AMOUNT (₹)</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>Basic Gross Monthly Wage</td>
                                <td style="text-align: right; font-weight: 600;">₹{base_salary:,.2f}</td>
                            </tr>
                            <tr>
                                <td>Sundays & Public Holidays ({total_paid_holidays} days)</td>
                                <td style="text-align: right; color: #059669; font-weight: 600;">Paid (0.0 Unpaid)</td>
                            </tr>
                            <tr>
                                <td>Assumed Working Days ({len(df_attendance[df_attendance['Status'].str.contains('Assumed')])} days)</td>
                                <td style="text-align: right; color: #059669; font-weight: 600;">Paid (0.0 Unpaid)</td>
                            </tr>
                            <tr class="total-row">
                                <td>GROSS EARNINGS (A)</td>
                                <td style="text-align: right; color: #0F172A;">₹{base_salary:,.2f}</td>
                            </tr>
                        </tbody>
                    </table>
                </div>

                <!-- Deductions Column -->
                <div>
                    <table>
                        <thead>
                            <tr>
                                <th>DEDUCTION HEAD</th>
                                <th style="text-align: right;">AMOUNT (₹)</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>Attendance LOP ({total_unpaid_equivalent:.1f} days)</td>
                                <td style="text-align: right; color: #DC2626; font-weight: 600;">₹{attendance_deduction:,.2f}</td>
                            </tr>
                            <tr>
                                <td>Cash Advances Recovered</td>
                                <td style="text-align: right; color: #DC2626; font-weight: 600;">₹{total_advances:,.2f}</td>
                            </tr>
                            <tr>
                                <td>Staff Store Purchases (Bills)</td>
                                <td style="text-align: right; color: #DC2626; font-weight: 600;">₹{total_staff_sales:,.2f}</td>
                            </tr>
                            <tr class="total-row">
                                <td>TOTAL DEDUCTIONS (B)</td>
                                <td style="text-align: right; color: #DC2626;">₹{(attendance_deduction + total_ledger_deductions):,.2f}</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>

            <div class="net-banner">
                <div>
                    <div class="net-lbl">FINAL NET SETTLEMENT DUE (A - B)</div>
                    <div class="net-words">In Words: {inr_words}</div>
                </div>
                <div class="net-amt">₹{net_payable:,.2f}</div>
            </div>

            <div class="sig-row">
                <div class="sig-box">PREPARED BY (ACCOUNTS)</div>
                <div class="sig-box">VERIFIED BY (AUDITOR / MGR)</div>
                <div class="sig-box">EMPLOYEE SIGNATURE</div>
            </div>
        </div>
    </body>
    </html>
    """

    components.html(voucher_full_html, height=730, scrolling=True)

    st.download_button(
        label="📥 Download Voucher HTML",
        data=voucher_full_html,
        file_name=f"Salary_Voucher_{clean_emp_display}_{month_name}_{active_year}.html",
        mime="text/html"
    )

# -------------------------------------------------------------
# TAB 2: ATTENDANCE CALENDAR & PUNCH AUDIT
# -------------------------------------------------------------
with tab2:
    st.markdown("### 🗓️ **Month Attendance & Punch Audit**")
    st.caption("Visual month calendar with punch times, daily sales performance, and condition-based status.")

    b1, b2, b3, b4, b5 = st.columns(5)
    present_full = len(df_attendance[df_attendance['Status'] == 'Full Day / Present'])
    half_days = len(df_attendance[df_attendance['Status'] == 'Half Day'])
    absent_days = len(df_attendance[df_attendance['Status'] == 'Leave / Absent'])
    holidays_count = len(df_attendance[df_attendance['Status'].str.contains('Paid Holiday')])
    assumed_days = len(df_attendance[df_attendance['Status'].str.contains('Assumed')])

    b1.metric("🟢 Full Days Worked", f"{present_full} Days")
    b2.metric("🟡 Half Days (<= 2 PM)", f"{half_days} Days")
    b3.metric("🔴 Absences / Leaves", f"{absent_days} Days")
    b4.metric("🔵 Paid Holidays", f"{holidays_count} Days")
    b5.metric("🟣 Assumed Future", f"{assumed_days} Days")

    st.markdown("<br/>", unsafe_allow_html=True)
    st.markdown("#### 📅 **Visual Monthly Calendar View**")

    # Build standalone HTML for the Calendar Grid
    cal_grid_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@500;600;700;800&family=JetBrains+Mono:wght@500;700&display=swap" rel="stylesheet">
    <style>
        body {{
            font-family: 'Plus Jakarta Sans', system-ui, sans-serif;
            margin: 0;
            padding: 8px;
            background: transparent;
        }}
        .cal-wrapper {{
            max-width: 100%;
            margin: 0 auto;
        }}
        .cal-header-row {{
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            gap: 8px;
            margin-bottom: 8px;
        }}
        .header-cell {{
            text-align: center;
            font-weight: 700;
            font-size: 12px;
            color: #475569;
            padding: 7px 0;
            background: #F1F5F9;
            border-radius: 6px;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }}
        .header-sun {{
            color: #DC2626;
            background: #FEE2E2;
        }}
        .cal-body-grid {{
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            gap: 8px;
        }}
        .day-card {{
            background: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 10px;
            padding: 8px 9px;
            min-height: 82px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            box-shadow: 0 1px 3px rgba(0,0,0,0.03);
            transition: all 0.15s ease;
        }}
        .day-card:hover {{
            border-color: #94A3B8;
            box-shadow: 0 4px 6px -1px rgba(0,0,0,0.08);
            transform: translateY(-1px);
        }}
        .day-num {{
            font-size: 14px;
            font-weight: 800;
            color: #0F172A;
        }}
        .badge {{
            font-size: 9.5px;
            font-weight: 700;
            padding: 2px 6px;
            border-radius: 9999px;
            display: inline-block;
            white-space: nowrap;
        }}
        .punch-text {{
            font-size: 10px;
            color: #475569;
            margin-top: 3px;
            font-family: 'JetBrains Mono', monospace;
        }}
        .sales-text {{
            font-size: 10.5px;
            font-weight: 700;
            color: #059669;
            margin-top: 2px;
        }}
    </style>
    </head>
    <body>
        <div class="cal-wrapper">
            <div class="cal-header-row">
                <div class="header-cell">MON</div>
                <div class="header-cell">TUE</div>
                <div class="header-cell">WED</div>
                <div class="header-cell">THU</div>
                <div class="header-cell">FRI</div>
                <div class="header-cell">SAT</div>
                <div class="header-cell header-sun">SUN</div>
            </div>
            <div class="cal-body-grid">
    """

    cal = calendar.Calendar(firstweekday=0)
    month_weeks = cal.monthdayscalendar(active_year, active_month)

    for week in month_weeks:
        for day_val in week:
            if day_val == 0:
                cal_grid_html += '<div style="background: #F8FAFC; border: 1px dashed #CBD5E1; border-radius: 10px; min-height: 82px; opacity: 0.35;"></div>\n'
            else:
                row_match = df_attendance[df_attendance['Day_Num'] == day_val].iloc[0]
                d_num = day_val
                status_short = row_match['Status'].split('(')[0].strip()
                b_bg = row_match['Badge_BG']
                b_color = row_match['Badge_Color']
                punch_in = row_match['Earliest Punch']
                punch_out = row_match['Latest Punch']
                sales_val = row_match['Total Sales (₹)']
                b_cnt = row_match['Bills Count']

                punch_html = ""
                if punch_in != "-":
                    punch_html = f'<div class="punch-text">🕒 {punch_in} - {punch_out}</div>'

                sales_html = ""
                if sales_val > 0:
                    sales_html = f'<div class="sales-text">₹{sales_val:,.0f} ({b_cnt} bills)</div>'

                cal_grid_html += f"""
                <div class="day-card">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <span class="day-num">{d_num:02d}</span>
                        <span class="badge" style="background:{b_bg}; color:{b_color};">{status_short}</span>
                    </div>
                    <div>
                        {punch_html}
                        {sales_html}
                    </div>
                </div>
                """

    cal_grid_html += """
            </div>
        </div>
    </body>
    </html>
    """

    components.html(cal_grid_html, height=540, scrolling=True)

    st.markdown("---")
    st.markdown("#### 📋 **Searchable Attendance Log & Punches**")

    f1, f2 = st.columns([2, 1])
    with f1:
        st_opts = list(df_attendance['Status'].unique())
        selected_sts = st.multiselect("Filter by Attendance Status:", st_opts, default=st_opts)

    filtered_log = df_attendance[df_attendance['Status'].isin(selected_sts)].copy()

    st.dataframe(
        filtered_log[['Date', 'Day_Name', 'Status', 'Unpaid Eq', 'Earliest Punch', 'Latest Punch', 'Bills Count', 'Total Sales (₹)']].style.format({
            'Total Sales (₹)': "₹{:,.2f}",
            'Unpaid Eq': "{:.1f}",
            'Date': lambda d: d.strftime('%d-%m-%Y') if hasattr(d, 'strftime') else str(d)
        }),
        use_container_width=True,
        hide_index=True
    )

    csv_att = df_attendance.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Complete Attendance Log (CSV)",
        data=csv_att,
        file_name=f"Attendance_Log_{clean_emp_display}_{month_name}_{active_year}.csv",
        mime="text/csv"
    )

# -------------------------------------------------------------
# TAB 3: LEDGER & ADVANCES AUDIT
# -------------------------------------------------------------
with tab3:
    st.markdown("### 💳 **Ledger Deductions & Cash Advances Audit**")
    st.caption("Complete breakdown of cash advances and staff personal purchases debited to the employee account.")

    l1, l2, l3 = st.columns(3)
    l1.metric("Cash Advances Recovered", f"₹{total_advances:,.2f}", f"{len(df_deductions[df_deductions['Category'] == 'Advance'])} Advance Entries")
    l2.metric("Staff Store Purchases", f"₹{total_staff_sales:,.2f}", f"{len(df_deductions[df_deductions['Category'] == 'Staff Sales'])} Store Bills")
    l3.metric("Total Ledger Deductions", f"₹{total_ledger_deductions:,.2f}", "Recovered from Monthly Salary", delta_color="inverse")

    st.markdown("<br/>", unsafe_allow_html=True)

    if df_deductions.empty:
        st.info("No ledger deductions (advances or staff sales) found for this period.")
    else:
        st.markdown("#### 🧾 **Itemized Deductions Table**")
        st.dataframe(
            df_deductions[['Date', 'Doc No', 'Type', 'Category', 'Remarks', 'Amount (₹)']].style.format({
                'Amount (₹)': "₹{:,.2f}",
                'Date': lambda d: d.strftime('%d-%m-%Y') if hasattr(d, 'strftime') else str(d)
            }),
            use_container_width=True,
            hide_index=True
        )

        csv_ded = df_deductions.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download Ledger Deductions (CSV)",
            data=csv_ded,
            file_name=f"Ledger_Deductions_{clean_emp_display}_{month_name}_{active_year}.csv",
            mime="text/csv"
        )

    if ignored_entries:
        with st.expander(f"ℹ️ Excluded Ledger Entries ({len(ignored_entries)} Ignored Records)", expanded=False):
            st.caption("Entries containing 'SALARY CASH' or 'Opening Balance' are ignored per company deduction guidelines.")
            df_ignored = pd.DataFrame(ignored_entries)
            st.dataframe(
                df_ignored.style.format({
                    'Amount': "₹{:,.2f}",
                    'Date': lambda d: d.strftime('%d-%m-%Y') if hasattr(d, 'strftime') else str(d)
                }),
                use_container_width=True,
                hide_index=True
            )

# -------------------------------------------------------------
# TAB 4: SALES & STORE PERFORMANCE
# -------------------------------------------------------------
with tab4:
    st.markdown("### 📈 **Staff Sales Performance & Revenue Intelligence**")
    st.caption("Store turnover attributed directly to the employee during active register hours.")

    tot_rev = df_attendance['Total Sales (₹)'].sum()
    tot_bills = df_attendance['Bills Count'].sum()
    active_days_count = len(df_attendance[df_attendance['Bills Count'] > 0])
    avg_daily_rev = tot_rev / active_days_count if active_days_count > 0 else 0
    avg_ticket = tot_rev / tot_bills if tot_bills > 0 else 0
    cost_to_rev_ratio = (base_salary / tot_rev * 100) if tot_rev > 0 else 0

    s1, s2, s3, s4 = st.columns(4)
    s1.metric("Total Revenue Attributed", f"₹{tot_rev:,.2f}", f"{active_days_count} Active Selling Days")
    s2.metric("Total Invoices Punched", f"{tot_bills:,} Bills", f"Avg {tot_bills/active_days_count:.1f} Bills/Day" if active_days_count else "0")
    s3.metric("Average Daily Revenue", f"₹{avg_daily_rev:,.2f}", "On Active Store Days")
    s4.metric("Avg Order Value (AOV)", f"₹{avg_ticket:,.2f}", f"Staff Cost: {cost_to_rev_ratio:.2f}% of Revenue")

    st.markdown("<br/>", unsafe_allow_html=True)
    st.markdown("#### 📊 **Daily Sales Trend ('Bill Amount' & Ticket Counts)**")

    chart_df = df_attendance[['Date', 'Total Sales (₹)', 'Bills Count']].copy()
    chart_df['Date_str'] = pd.to_datetime(chart_df['Date']).dt.strftime('%d %b')

    base = alt.Chart(chart_df).encode(
        x=alt.X('Date_str:N', title='Date of Month', sort=None)
    )

    area = base.mark_area(
        opacity=0.22,
        color="#2563EB"
    ).encode(
        y=alt.Y('Total Sales (₹):Q', title='Daily Sales Value (₹)', axis=alt.Axis(format=',.0f'))
    )

    line = base.mark_line(
        color="#1D4ED8",
        strokeWidth=3,
        point=alt.OverlayMarkDef(filled=True, size=75, color="#1D4ED8")
    ).encode(
        y=alt.Y('Total Sales (₹):Q'),
        tooltip=[
            alt.Tooltip('Date_str:N', title='Date'),
            alt.Tooltip('Total Sales (₹):Q', title='Sales Amount (₹)', format=',.2f'),
            alt.Tooltip('Bills Count:Q', title='Bills Generated')
        ]
    )

    combined_chart = (area + line).properties(
        height=360
    ).interactive()

    st.altair_chart(combined_chart, use_container_width=True)

# -------------------------------------------------------------
# TAB 5: BUSINESS RULES & VALIDATION AUDIT
# -------------------------------------------------------------
with tab5:
    st.markdown("### 🔍 **Compliance & Audit Rule Verification**")
    st.caption("Verification trail for identity cross-matching, Kerala holiday calendar, and payroll formulas.")

    st.markdown(f"""
    #### **1. Identity Validation Trail**
    * **Sales Register Extracted User:** `{name_a}`
    * **Ledger Extracted Account:** `{name_b}`
    * **Normalized Sales Register:** `{norm_a}`
    * **Normalized Ledger Account:** `{norm_b}`
    * **Validation Outcome:** `MATCHED ✓` (Case-insensitive comparison ignoring `"FT "` prefix)

    ---

    #### **2. Calendar & Working Days Formulation**
    * **Active Period:** `{month_name} {active_year}`
    * **Total Calendar Days in Month:** `{total_days_in_month}` days
    * **Recognized Public Holidays (Kerala, IN):**
      - `{", ".join([f"{d.strftime('%d-%b-%Y')} ({n})" for d, n in kl_holidays.items() if d.month == active_month and d.year == active_year]) or "None"}`
    * **Total Sundays in Month:** `{len([d for d in paid_holiday_dates if d.weekday() == 6])}` Sundays
    * **Total Distinct Paid Holidays:** `{total_paid_holidays}` days
    * **Net Working Days:** `{total_days_in_month} - {total_paid_holidays} = {working_days}` days
    * **Daily Wage Formula:** `₹{base_salary:,.2f} / {working_days} = ₹{daily_wage:,.2f} / day`

    ---

    #### **3. Attendance Condition Logic**
    * **Condition A (Paid Holiday):** Sundays & Kerala Govt holidays marked as Paid Holiday (`0.0` unpaid eq).
    * **Condition B (Worked Day):**
      - Last punch at or before `14:00` (2:00 PM) $\\rightarrow$ Half Day (`0.5` unpaid eq).
      - Last punch after `14:00` $\\rightarrow$ Full Day / Present (`0.0` unpaid eq).
    * **Condition C (Incomplete Future Data):** Dates strictly after `{max_date_in_file.strftime('%d-%b-%Y')}` marked as Present (Assumed) (`0.0` unpaid eq).
    * **Condition D (Absent):** Dates on or before `{max_date_in_file.strftime('%d-%b-%Y')}` with no sales activity marked as Leave / Absent (`1.0` unpaid eq).
    * **Total Unpaid Absence Days:** `{total_unpaid_equivalent:.1f}` days
    * **Attendance Deduction:** `{total_unpaid_equivalent:.1f} × ₹{daily_wage:,.2f} = ₹{attendance_deduction:,.2f}`

    ---

    #### **4. Ledger Deduction Logic**
    * **Advances:** Debited entries with Type `Payment` and Remarks containing `ADVANCE CASH` $\\rightarrow$ `₹{total_advances:,.2f}`.
    * **Staff Sales:** Debited entries with Type `Sales` $\\rightarrow$ `₹{total_staff_sales:,.2f}`.
    * **Ignored Entries:** Entries with `SALARY CASH` or `Opening Balance` are filtered out.
    * **Total Ledger Deductions:** `₹{total_ledger_deductions:,.2f}`

    ---

    #### **5. Final Settlement Math**
    $$\\text{{Net Payable}} = \\text{{Base Salary}} - \\text{{Attendance Deduction}} - \\text{{Advances}} - \\text{{Staff Sales}}$$
    $$\\text{{Net Payable}} = ₹{base_salary:,.2f} - ₹{attendance_deduction:,.2f} - ₹{total_advances:,.2f} - ₹{total_staff_sales:,.2f} = \\mathbf{{₹{net_payable:,.2f}}}$$
    """)
