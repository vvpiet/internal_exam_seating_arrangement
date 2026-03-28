import streamlit as st
import pandas as pd
from datetime import date, timedelta
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from fpdf import FPDF
import io
import os

st.set_page_config(page_title="Exam Supervision Chart", layout="wide")

# Custom CSS for styling
st.markdown("""
    <style>
    body {
        background-color: #f5f5f5;
    }
    .header-container {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 20px;
        padding: 20px;
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        border-radius: 10px;
        margin-bottom: 30px;
        color: white;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .header-text h1 {
        margin: 0;
        font-size: 32px;
        font-weight: bold;
        text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.3);
    }
    .footer-container {
        margin-top: 50px;
        padding: 20px;
        background: linear-gradient(135deg, #2a5298 0%, #1e3c72 100%);
        border-radius: 10px;
        color: white;
        text-align: center;
        font-size: 14px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .footer-text {
        margin: 10px 0;
        font-weight: 500;
    }
    </style>
""", unsafe_allow_html=True)

# Header with Logo and Title
col1, col2 = st.columns([1, 4])
with col1:
    # Try to load logo if it exists, otherwise use a placeholder icon
    logo_path = "vvp_logo.png"
    if os.path.exists(logo_path):
        st.image(logo_path, width=100)
    else:
        # Using a professional university icon as placeholder
        st.image("https://img.icons8.com/color/200/000000/organization.png", width=100)
        st.caption("(Add vvp_logo.png to project folder)")
with col2:
    st.markdown("""
        <div style='color: #1e3c72; padding: 20px;'>
            <h1>VVP Institute of Engineering and Technology, Solapur</h1>
            <h3 style='color: #2a5298; margin-top: -10px;'>Internal Exam Supervision Chart Generator</h3>
        </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# Sidebar Inputs
st.sidebar.header("Configuration")
st.sidebar.info("💡 To use your institute logo, save it as 'vvp_logo.png' in the project folder")

# Date Range Input
col1, col2 = st.sidebar.columns(2)
with col1:
    exam_start_date = st.date_input("Exam Start Date")
with col2:
    exam_end_date = st.date_input("Exam End Date")

# Holidays Input
st.sidebar.subheader("Holidays to Exclude")
holidays_input = st.sidebar.text_area(
    "Enter holidays (one date per line, format: DD-MM-YYYY)",
    "25-12-2024\n26-12-2024"
)

holidays = []
if holidays_input.strip():
    for holiday in holidays_input.strip().split("\n"):
        try:
            holiday_date = pd.to_datetime(holiday.strip(), format="%d-%m-%Y").date()
            holidays.append(holiday_date)
        except:
            st.sidebar.warning(f"Invalid date format: {holiday.strip()}")

st.sidebar.info(f"Excluded holidays: {len(holidays)} dates")

# Blocks per Day Configuration
st.sidebar.subheader("Blocks Per Day")
col1, col2 = st.sidebar.columns(2)
with col1:
    morning_blocks = st.number_input("Morning Blocks", min_value=1, value=2, step=1)
with col2:
    evening_blocks = st.number_input("Evening Blocks", min_value=0, value=2, step=1)

# Morning Time Slots
st.sidebar.write("**Morning Time Slots (comma separated)**")
morning_slots_input = st.sidebar.text_input(
    "Morning Slots",
    "11:00-12:00,12:15-1:15"
)

# Evening Time Slots
st.sidebar.write("**Evening Time Slots (comma separated)**")
evening_slots_input = st.sidebar.text_input(
    "Evening Slots",
    "2:00-3:00,3:15-4:15"
)

morning_time_slots = [s.strip() for s in morning_slots_input.split(",") if s.strip()]
evening_time_slots = [s.strip() for s in evening_slots_input.split(",") if s.strip()]

# Combine time slots based on block configuration
time_slots = []
if morning_blocks > 0:
    time_slots.extend(morning_time_slots[:morning_blocks])
if evening_blocks > 0:
    time_slots.extend(evening_time_slots[:evening_blocks])

# Faculty List Import from Excel
st.sidebar.subheader("Faculty List")
excel_file = st.sidebar.file_uploader("Upload Faculty List (Excel)", type=["xlsx", "xls"])

faculty_list = []

if excel_file is not None:
    try:
        df_faculty = pd.read_excel(excel_file)
        # Assume faculty names are in the first column
        faculty_list = [f.strip() for f in df_faculty.iloc[:, 0].astype(str) if f.strip() and f.strip() != 'nan']
        st.sidebar.success(f"Loaded {len(faculty_list)} faculty members from Excel")
    except Exception as e:
        st.sidebar.error(f"Error reading Excel file: {e}")
else:
    st.sidebar.info("Please upload an Excel file with faculty names in the first column")
    
    # Fallback: Manual entry
    st.sidebar.write("**Or enter manually (comma separated)**")
    senior_faculty = st.sidebar.text_area(
        "Senior Faculty",
        "Dr. Patil,Dr. Mehta"
    ).split(",")

    junior_faculty = st.sidebar.text_area(
        "Junior Faculty",
        "Prof. Shah,Prof. Kumar,Prof. Rao"
    ).split(",")

    faculty_list = [f.strip() for f in senior_faculty] + [f.strip() for f in junior_faculty]
    faculty_list = [f for f in faculty_list if f]


if st.button("Generate Chart"):
    if not faculty_list:
        st.error("Please provide faculty list (either upload Excel or enter manually)")
    elif exam_start_date > exam_end_date:
        st.error("Exam start date must be before end date")
    elif len(time_slots) == 0:
        st.error("Please configure at least one time slot")
    else:
        # Generate date range excluding holidays
        current_date = exam_start_date
        date_range = []
        while current_date <= exam_end_date:
            if current_date not in holidays:
                date_range.append(current_date)
            current_date += timedelta(days=1)

        total_slots = len(date_range) * len(time_slots)
        num_faculty = len(faculty_list)
        
        # Determine how many faculty per slot needed to use all faculty
        faculty_per_slot = max(1, (num_faculty + total_slots - 1) // total_slots)  # Ceiling division
        
        data = []
        faculty_index = 0
        sr_no = 1
        
        # Create supervision chart - one row per date per supervisor per slot
        for day_idx, exam_date in enumerate(date_range):
            for slot_idx, slot in enumerate(time_slots):
                # Determine if slot is Morning (M) or Evening (E)
                is_morning = slot_idx < morning_blocks
                period = "M" if is_morning else "E"
                
                # Assign multiple faculty to this slot if needed
                for fp in range(faculty_per_slot):
                    if faculty_index < num_faculty:  # Don't exceed available faculty
                        assigned_faculty = faculty_list[faculty_index]
                        
                        # Build row with all time slots
                        row = {
                            "Sr. No.": sr_no,
                            "Supervisor Name": assigned_faculty,
                            "Date": exam_date.strftime("%d-%m-%Y"),
                            "M/E": period
                        }
                        
                        # Add tick marks for each time slot
                        for ts_idx, ts in enumerate(time_slots):
                            if ts_idx == slot_idx:
                                row[ts] = "✓"
                            else:
                                row[ts] = ""
                        
                        sr_no += 1
                        faculty_index += 1
                        data.append(row)

        df = pd.DataFrame(data)

        # Reorder columns: Sr. No., Supervisor Name, Date, M/E, then time slots
        column_order = ["Sr. No.", "Supervisor Name", "Date", "M/E"]
        column_order.extend(time_slots)
        df = df[column_order]

        st.subheader(f"Supervision Chart ({exam_start_date.strftime('%d-%m-%Y')} to {exam_end_date.strftime('%d-%m-%Y')})")
        total_exam_days = (exam_end_date - exam_start_date).days + 1
        st.write(f"**Total Days in Range:** {total_exam_days} | **Exam Days (after excluding {len(holidays)} holidays):** {len(date_range)} | **Time Slots/Day:** {len(time_slots)} | **Total Supervision Slots:** {total_slots}")
        st.write(f"**Total Faculty:** {num_faculty} | **Faculty per Slot:** {faculty_per_slot} | **Faculty Allocated:** {faculty_index}")
        st.dataframe(df, use_container_width=True)

        # CSV Download with Headers
        csv_buffer = io.StringIO()
        csv_buffer.write("VVP Institute of engineering and Technology, Solapur\n")
        csv_buffer.write("Internal Supervision Chart\n")
        csv_buffer.write(f"Period: {exam_start_date.strftime('%d-%m-%Y')} to {exam_end_date.strftime('%d-%m-%Y')}\n")
        csv_buffer.write(f"Morning Blocks: {morning_blocks} | Evening Blocks: {evening_blocks}\n")
        csv_buffer.write("\n")
        csv_buffer.write(df.to_csv(index=False, encoding='utf-8'))
        csv_content = csv_buffer.getvalue()
        
        st.download_button(
            label="Download CSV",
            data=csv_content,
            file_name="supervision_chart.csv",
            mime="text/csv"
        )

        # Word Download
        doc = Document()
        doc.add_heading("VVP Institute of engineering and Technology, Solapur", 0)
        doc.add_heading("Internal Supervision Chart", level=1)
        doc.add_paragraph(f"Period: {exam_start_date.strftime('%d-%m-%Y')} to {exam_end_date.strftime('%d-%m-%Y')}")
        doc.add_paragraph(f"Morning Blocks: {morning_blocks} | Evening Blocks: {evening_blocks}")
        doc.add_paragraph("")

        table = doc.add_table(rows=len(df)+1, cols=len(df.columns))
        table.style = 'Light Grid Accent 1'

        for j, col in enumerate(df.columns):
            table.rows[0].cells[j].text = col

        for i in range(len(df)):
            for j in range(len(df.columns)):
                table.rows[i+1].cells[j].text = str(df.iloc[i, j])

        # Save Word to bytes buffer
        doc_buffer = io.BytesIO()
        doc.save(doc_buffer)
        doc_buffer.seek(0)

        st.download_button(
            label="Download Word",
            data=doc_buffer.getvalue(),
            file_name="supervision_chart.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        # PDF Download
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", size=14)
        pdf.cell(0, 8, "VVP Institute of engineering and Technology, Solapur", ln=True, align="C")
        pdf.set_font("Arial", "B", size=12)
        pdf.cell(0, 8, "Internal Supervision Chart", ln=True, align="C")
        pdf.set_font("Arial", size=10)
        pdf.cell(0, 8, f"Period: {exam_start_date.strftime('%d-%m-%Y')} to {exam_end_date.strftime('%d-%m-%Y')}", ln=True)
        pdf.cell(0, 8, f"Morning Blocks: {morning_blocks} | Evening Blocks: {evening_blocks}", ln=True)
        pdf.ln(5)

        # Create a copy of dataframe with tick marks replaced for PDF compatibility
        df_pdf = df.copy()
        # Replace tick marks with 'X' for PDF export (latin-1 compatibility)
        for col in df_pdf.columns:
            df_pdf[col] = df_pdf[col].astype(str).str.replace('✓', 'X')

        # Add table headers
        col_width = pdf.w / len(df_pdf.columns)
        for col in df_pdf.columns:
            pdf.cell(col_width, 8, str(col)[:15], border=1, align="C")
        pdf.ln()

        # Add table data
        for i in range(len(df_pdf)):
            for j in range(len(df_pdf.columns)):
                cell_value = str(df_pdf.iloc[i, j])[:15]
                # Ensure all characters are latin-1 compatible
                try:
                    cell_value.encode('latin-1')
                except UnicodeEncodeError:
                    cell_value = '?'
                pdf.cell(col_width, 8, cell_value, border=1, align="C")
            pdf.ln()

        # Save PDF to bytes buffer - use output() without arguments to get bytes
        pdf_bytes = pdf.output()
        
        st.download_button(
            label="Download PDF",
            data=pdf_bytes,
            file_name="supervision_chart.pdf",
            mime="application/pdf"
        )

# Footer
st.markdown("---")
st.markdown("""
    <div class='footer-container'>
        <div class='footer-text'>Prepared by</div>
        <div style='font-size: 16px; font-weight: bold;'>Prof. Amir M. Usman Wagdarikar</div>
        <div style='font-size: 13px;'>Asst. Prof., Department of Electronics and Telecommunication</div>
        <div style='margin-top: 10px; font-size: 12px; opacity: 0.9;'>VVP Institute of Engineering and Technology, Solapur</div>
    </div>
""", unsafe_allow_html=True)