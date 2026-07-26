import io
import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from apscheduler.schedulers.background import BackgroundScheduler
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from reportlab.lib.pagesizes import letter
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table
import streamlit as st

# ==========================================
# 1. PAGE CONFIGURATION & STYLING
# ==========================================
st.set_page_config(
    page_title="Client Risk & Portfolio Analytics",
    page_icon="📈",
    layout="wide",
)

st.title("📈 Client Risk Assessment & Portfolio Analytics Dashboard")

# ==========================================
# 2. INITIALIZE BACKGROUND SCHEDULER
# ==========================================
if "scheduler" not in st.session_state:
    scheduler = BackgroundScheduler()
    scheduler.start()
    st.session_state.scheduler = scheduler


# ==========================================
# 3. HELPER FUNCTIONS (PDF & SCHEDULER JOB)
# ==========================================
def create_pdf_report(advisor_name, cdata, div_table):
    """Generates a PDF report buffer for a given advisor."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    story = []

    # Filter Advisor Data
    adv_clients = cdata[cdata["Advisor"] == advisor_name]

    # PDF Title
    story.append(
        Paragraph(
            f"<b>Advisor Portfolio Report: {advisor_name}</b>", style=None
        )
    )
    story.append(Spacer(1, 12))

    # Risk Profile Summary
    summary = adv_clients["Risk Profile"].value_counts().reset_index()
    summary.columns = ["Risk Profile", "Client Count"]

    data = [summary.columns.tolist()] + summary.values.tolist()
    t = Table(data)
    story.append(t)
    story.append(Spacer(1, 20))

    doc.build(story)
    buffer.seek(0)
    return buffer


def send_email_job(recipient_email, advisor_name):
    """Standalone background task triggered by APScheduler (No st.* calls inside!)."""
    try:
        # Load raw data inside background job context
        cdata = pd.read_excel("Client_Data.xlsx")
        pdata = pd.read_excel("Portfolio_Data.xlsx")

        def risk_profile(rt):
            try:
                rt = float(rt)
            except (TypeError, ValueError):
                return "Error"
            if rt <= 3:
                return "Conservative"
            elif rt <= 6:
                return "Balanced"
            else:
                return "Aggressive"

        cdata["Risk Profile"] = cdata["Risk Tolerance"].apply(risk_profile)

        # Portfolio Analytics Summary
        totals = (
            pdata.groupby("Client ID")["Return"]
            .sum()
            .reset_index()
            .rename(columns={"Return": "Total Return"})
        )
        df = pdata.merge(totals, on="Client ID", how="left")
        df["Percentage"] = (df["Return"] / df["Total Return"]) * 100
        pivot = pd.pivot_table(
            df,
            index="Client ID",
            columns="Asset Class",
            values="Percentage",
            aggfunc="sum",
            fill_value=0,
        )

        # Generate PDF
        pdf_buffer = create_pdf_report(advisor_name, cdata, pivot)

        # SMTP Email Credentials (Replace with your actual sender details)
        sender_email = "your_email@gmail.com"
        sender_password = "your_app_password"

        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = recipient_email
        msg["Subject"] = f"Weekly Portfolio Report - {advisor_name}"

        body = f"Hello,\n\nPlease find attached the weekly portfolio summary report for {advisor_name}.\n\nBest regards,\nAnalytics Team"
        msg.attach(MIMEText(body, "plain"))

        attachment = MIMEApplication(pdf_buffer.read(), _subtype="pdf")
        attachment.add_header(
            "Content-Disposition",
            "attachment",
            filename=f"{advisor_name}_Weekly_Report.pdf",
        )
        msg.attach(attachment)

        # Send via SMTP
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, sender_password)
            server.send_message(msg)

        print(
            f"[Scheduler Success] Report sent to {recipient_email} for advisor {advisor_name}."
        )

    except Exception as e:
        print(f"[Scheduler Error] Failed to send scheduled email: {e}")


# ==========================================
# 4. DATA LOADING & CLEANING
# ==========================================


@st.cache_data
def load_and_clean_data():
    cdata = pd.read_excel("Client_Data.xlsx")
    pdata = pd.read_excel("Portfolio_Data.xlsx")

    def risk_profile(rt):
        try:
            rt = float(rt)
        except (TypeError, ValueError):
            return "Error"
        if rt <= 3:
            return "Conservative"
        elif rt <= 6:
            return "Balanced"
        else:
            return "Aggressive"

    cdata["Risk Profile"] = cdata["Risk Tolerance"].apply(risk_profile)

    if "Date" not in pdata.columns:
        pdata["Date"] = pd.date_range(
            start="2024-01-01", periods=len(pdata), freq="D"
        )
    else:
        pdata["Date"] = pd.to_datetime(pdata["Date"])

    if "Portfolio Value" not in pdata.columns:
        pdata["Portfolio Value"] = pdata.get("Return", 1000) * 10

    return cdata, pdata


cdata, pdata = load_and_clean_data()

# ==========================================
# 5. SIDEBAR FILTERS
# ==========================================
st.sidebar.header("🔍 Filters")

min_date = pdata["Date"].min().date()
max_date = pdata["Date"].max().date()

start_date, end_date = st.sidebar.date_input(
    "Select Date Range",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date,
)

advisors = ["All"] + list(cdata["Advisor"].dropna().unique())
selected_advisor = st.sidebar.selectbox("Select Advisor", advisors)

filtered_pdata = pdata[
    (pdata["Date"].dt.date >= start_date) & (pdata["Date"].dt.date <= end_date)
]

if selected_advisor != "All":
    filtered_cdata = cdata[cdata["Advisor"] == selected_advisor]
    filtered_pdata = filtered_pdata[
        filtered_pdata["Client ID"].isin(filtered_cdata["Client ID"])
    ]
else:
    filtered_cdata = cdata.copy()

# ==========================================
# 6. CORE COMPUTATION & ANALYTICS
# ==========================================


def analyze_portfolio(pdata):
    totals = (
        pdata.groupby("Client ID")["Return"]
        .sum()
        .reset_index()
        .rename(columns={"Return": "Total Return"})
    )
    df = pdata.merge(totals, on="Client ID", how="left")
    df["Percentage"] = (df["Return"] / df["Total Return"]) * 100

    pivot = pd.pivot_table(
        df,
        index="Client ID",
        columns="Asset Class",
        values="Percentage",
        aggfunc="sum",
        fill_value=0,
    )

    asset_counts = (pivot > 0).sum(axis=1)
    max_assets = max(len(pivot.columns), 1)
    pivot["Diversification Score"] = (asset_counts / max_assets) * 100

    return pivot, df


div_table, df_merged = analyze_portfolio(filtered_pdata)

# ==========================================
# 7. DASHBOARD TABS
# ==========================================
tab1, tab2, tab3, tab4 = st.tabs(
    [
        "📊 Risk Analytics",
        "💼 Portfolio Analytics",
        "👨‍💼 Advisor Analytics",
        "📄 PDF Reports & Automation",
    ]
)

# ------------------------------------------
# TAB 1: RISK ANALYTICS
# ------------------------------------------
with tab1:
    st.header("Risk Profile Breakdown")
    col1, col2 = st.columns([1, 1])

    with col1:
        risk_counts = filtered_cdata["Risk Profile"].value_counts().reset_index()
        risk_counts.columns = ["Risk Profile", "Count"]

        fig_risk = px.pie(
            risk_counts,
            names="Risk Profile",
            values="Count",
            title="Conservative / Balanced / Aggressive Distribution",
            color="Risk Profile",
            color_discrete_map={
                "Conservative": "#2ca02c",
                "Balanced": "#ff7f0e",
                "Aggressive": "#d62728",
            },
            hole=0.4,
        )
        st.plotly_chart(fig_risk, use_container_width=True)

    with col2:
        st.subheader("⚠️ High-Risk Clients Needing Review")
        high_risk_clients = filtered_cdata[
            filtered_cdata["Risk Profile"] == "Aggressive"
        ]

        if not high_risk_clients.empty:
            st.dataframe(
                high_risk_clients[
                    ["Client ID", "First Name", "Last Name", "Advisor"]
                ],
                use_container_width=True,
            )
        else:
            st.success("No high-risk clients found for current selection.")

# ------------------------------------------
# TAB 2: PORTFOLIO ANALYTICS
# ------------------------------------------
with tab2:
    st.header("Portfolio Breakdown")

    m1, m2, m3 = st.columns(3)
    top_asset = (
        filtered_pdata.groupby("Asset Class")["Return"].sum().idxmax()
        if not filtered_pdata.empty
        else "N/A"
    )
    avg_div_score = (
        div_table["Diversification Score"].mean()
        if not div_table.empty
        else 0
    )

    m1.metric("Top-Performing Asset Class", top_asset)
    m2.metric("Avg Diversification Score", f"{avg_div_score:.1f} / 100")
    m3.metric("Total Returns Logged", f"${filtered_pdata['Return'].sum():,.2f}")

    st.markdown("---")
    col1, col2 = st.columns([1, 1])

    with col1:
        asset_alloc = (
            filtered_pdata.groupby("Asset Class")["Return"].sum().reset_index()
        )
        fig_asset = px.bar(
            asset_alloc,
            x="Asset Class",
            y="Return",
            color="Asset Class",
            title="Overall Asset Allocation (Returns)",
        )
        st.plotly_chart(fig_asset, use_container_width=True)

    with col2:
        fig_div = px.histogram(
            div_table,
            x="Diversification Score",
            nbins=10,
            title="Diversification Score Distribution",
            color_discrete_sequence=["#1f77b4"],
        )
        st.plotly_chart(fig_div, use_container_width=True)

# ------------------------------------------
# TAB 3: ADVISOR ANALYTICS
# ------------------------------------------
with tab3:
    st.header("Advisor Performance Overview")

    adv_summary = (
        cdata.groupby("Advisor")
        .agg(
            Clients_Count=("Client ID", "count"),
            Avg_Risk_Tolerance=("Risk Tolerance", "mean"),
        )
        .reset_index()
    )

    aum_df = filtered_pdata.merge(cdata, on="Client ID")
    aum_summary = (
        aum_df.groupby("Advisor")["Portfolio Value"]
        .sum()
        .reset_index()
        .rename(columns={"Portfolio Value": "Total AUM"})
    )

    adv_metrics = adv_summary.merge(aum_summary, on="Advisor", how="left").fillna(
        0
    )

    col1, col2 = st.columns([1, 1])

    with col1:
        fig_aum = px.bar(
            adv_metrics,
            x="Advisor",
            y="Total AUM",
            title="Assets Under Management (AUM) per Advisor",
            text_auto=".2s",
        )
        st.plotly_chart(fig_aum, use_container_width=True)

    with col2:
        fig_risk_adv = px.scatter(
            adv_metrics,
            x="Clients_Count",
            y="Avg_Risk_Tolerance",
            size="Total AUM",
            color="Advisor",
            title="Clients vs Avg Risk Tolerance (Bubble size = AUM)",
        )
        st.plotly_chart(fig_risk_adv, use_container_width=True)

    st.subheader("Advisor Details")
    st.dataframe(adv_metrics, use_container_width=True)

# ------------------------------------------
# TAB 4: PDF REPORTS & AUTOMATION
# ------------------------------------------
with tab4:
    st.header("📄 Generate & Schedule PDF Reports")

    # On-demand PDF Download
    selected_pdf_advisor = st.selectbox(
        "Select Advisor for Instant PDF Download", cdata["Advisor"].unique()
    )

    if st.button("Generate Advisor PDF Report"):
        pdf_buffer = create_pdf_report(selected_pdf_advisor, cdata, div_table)
        st.download_button(
            label="📥 Download PDF Report",
            data=pdf_buffer,
            file_name=f"{selected_pdf_advisor}_Report.pdf",
            mime="application/pdf",
        )

    st.markdown("---")
    st.subheader("⏰ Weekly Report Email Scheduler")

    email_recipient = st.text_input(
        "Advisor Email Address", value="advisor@example.com"
    )
    scheduled_advisor = st.selectbox(
        "Select Advisor for Scheduled Report", cdata["Advisor"].unique()
    )

    # Convert readable day name to cron string
    day_mapping = {
        "Monday": "mon",
        "Tuesday": "tue",
        "Wednesday": "wed",
        "Thursday": "thu",
        "Friday": "fri",
        "Saturday": "sat",
        "Sunday": "sun",
    }

    scheduled_day_str = st.selectbox(
        "Select Delivery Day", list(day_mapping.keys())
    )
    scheduled_hour = st.slider("Select Delivery Hour (24-hr format)", 0, 23, 9)

    if st.button("Schedule Weekly Reports"):
        job_id = f"weekly_job_{scheduled_advisor}"

        # Overwrite existing job if already present
        if st.session_state.scheduler.get_job(job_id):
            st.session_state.scheduler.remove_job(job_id)

        # Add cron job using APScheduler
        st.session_state.scheduler.add_job(
            send_email_job,
            "cron",
            day_of_week=day_mapping[scheduled_day_str],
            hour=scheduled_hour,
            minute=0,
            args=[email_recipient, scheduled_advisor],
            id=job_id,
        )

        st.success(
            f"Successfully scheduled weekly report for **{scheduled_advisor}**!\n"
            f"Will send to **{email_recipient}** every **{scheduled_day_str}** at **{scheduled_hour}:00**."
        )

    # Active Scheduled Jobs List
    st.markdown("### 📋 Active Scheduled Jobs")
    jobs = st.session_state.scheduler.get_jobs()
    if jobs:
        for job in jobs:
            st.write(
                f"📌 **Job ID:** `{job.id}` | **Next Run Time:** {job.next_run_time}"
            )
    else:
        st.info("No active scheduled jobs currently registered.")