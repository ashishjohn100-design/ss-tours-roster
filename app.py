import streamlit as st
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

# Page Configuration
st.set_page_config(page_title="SS Tours Cab Roster", page_icon="🚐", layout="centered")

st.title("🚐 SS Tours Dehradun - Cab Roster Generator")
st.markdown("Easily generate daily rosters, update your Excel master tracker, and copy WhatsApp messages instantly!")

# Pre-loaded ground staff list
staff_list = [
    "PRABHA",
    "ASHMITA CHHETRI",
    "NEHA PUNDIR",
    "SADAF ABBASI",
    "LALITA",
    "JYOTI",
    "AYUSHI SINGH",
    "ABHISHEK KUMAR",
    "AARIYS CHAUHAN",
    "HARSH SHARMA"
]

# Date Input
default_date = datetime.now() + timedelta(days=1)
selected_date = st.date_input("Select Schedule Date", value=default_date)
roster_date_str = selected_date.strftime("%d %B %Y").upper()
sheet_name_date = selected_date.strftime("%d_%b_%Y").upper()

st.markdown("---")
st.subheader("Cab Assignments & Timings")

cabs_data = {}
cab_names = ["CAB 1", "CAB 2", "CAB 3"]

for cab in cab_names:
    st.markdown(f"### {cab}")
    col1, col2 = st.columns([2, 1])
    
    with col1:
        selected_members = st.multiselect(
            f"Select staff for {cab}",
            options=staff_list,
            key=f"staff_{cab}"
        )
    with col2:
        cab_time = st.text_input(f"Pickup Time ({cab})", placeholder="e.g., 4:55 am", key=f"time_{cab}")
        
    cabs_data[cab] = {"names": selected_members, "time": cab_time}

st.markdown("---")

if st.button("🚀 Generate Roster, Excel Update & Notepad Backup", type="primary"):
    
    # Check for duplicate assignments across cabs
    all_selected = []
    has_duplicate = False
    for cab in cab_names:
        for name in cabs_data[cab]["names"]:
            if name in all_selected:
                has_duplicate = True
                st.error(f"⚠️ Error: '{name}' is assigned to multiple cabs! Please make sure each staff member is in only one cab.")
            all_selected.append(name)
            
    if not has_duplicate:
        # --- Generate WhatsApp Format ---
        wa_output = f"{roster_date_str}\nCAB PICKUP SCHEDULE\nPLEASE FIND THE PICKUP SCHEDULE\n"
        
        for cab_name in cab_names:
            wa_output += f"\n{cab_name}\n"
            members = cabs_data[cab_name]["names"]
            time_str = cabs_data[cab_name]["time"]
            
            if members:
                for idx, member_name in enumerate(members, 1):
                    if idx == 1 and time_str:
                        wa_output += f"{idx}.\t{member_name} \t {time_str}\n"
                    else:
                        wa_output += f"{idx}.\t{member_name}\n"
            else:
                wa_output += "1.\t-\n"
                
        wa_output += ("\nIMPORTANT NOTICE:\n"
                      "•\tIF ANYONE IS LEFT OUT OF THE PICKUP LIST, KINDLY SHARE YOUR NAME.\n"
                      "•\tPLEASE ALSO CONFIRM IF YOU WILL NOT BE AVAILING THE PICKUP OR DROP SERVICE.\n"
                      "•\tTRAVEL SAFE AND HAVE A WONDERFUL DAY AHEAD WITH SS TOURS DEHRADUN")

        # --- Save to Notepad (.txt) ---
        txt_filename = f"ROSTER_{roster_date_str.replace(' ', '_')}.txt"
        with open(txt_filename, "w", encoding="utf-8") as f:
            f.write(wa_output)

        # --- Update Excel Multi-tab File ---
        file_name = "SS_Tours_Cab_Roster_Tracker.xlsx"
        try:
            if os.path.exists(file_name):
                wb = openpyxl.load_workbook(file_name)
            else:
                wb = openpyxl.Workbook()
                if "Sheet" in wb.sheetnames:
                    wb.remove(wb["Sheet"])

            base_sheet_name = sheet_name_date
            counter = 1
            while sheet_name_date in wb.sheetnames:
                sheet_name_date = f"{base_sheet_name}_{counter}"
                counter += 1

            ws = wb.create_sheet(title=sheet_name_date)
            ws.views.sheetView[0].showGridLines = True

            header_fill = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            zebra_fill = PatternFill(start_color="F2F4F4", end_color="F2F4F4", fill_type="solid")
            white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
            thin_border = Border(
                left=Side(style='thin', color='D5D8DC'), right=Side(style='thin', color='D5D8DC'),
                top=Side(style='thin', color='D5D8DC'), bottom=Side(style='thin', color='D5D8DC')
            )

            ws.merge_cells('A1:E1')
            title_cell = ws['A1']
            title_cell.value = f"SS TOURS DEHRADUN - CAB PICKUP SCHEDULE ({roster_date_str})"
            title_cell.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
            title_cell.fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
            title_cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 30

            headers = ["Date", "Cab Name", "Staff Name", "Pickup Time", "Status / Remarks"]
            ws.append([]) 
            ws.append(headers)
            ws.row_dimensions[3].height = 25

            for col_num, header in enumerate(headers, 1):
                cell = ws.cell(row=3, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border

            row_idx = 4
            for cab_name in cab_names:
                members = cabs_data[cab_name]["names"]
                time_str = cabs_data[cab_name]["time"]
                
                if members:
                    for idx, member_name in enumerate(members, 1):
                        pickup_t = time_str if idx == 1 else ""
                        row_data = [roster_date_str, cab_name, member_name, pickup_t, "Scheduled"]
                        ws.append(row_data)
                        ws.row_dimensions[row_idx].height = 20
                        
                        current_fill = zebra_fill if (row_idx % 2 == 0) else white_fill
                        for col_idx in range(1, 6):
                            c = ws.cell(row=row_idx, column=col_idx)
                            c.fill = current_fill
                            c.font = Font(name="Calibri", size=11)
                            c.border = thin_border
                            if col_idx in [1, 2, 4, 5]:
                                c.alignment = Alignment(horizontal="center", vertical="center")
                            else:
                                c.alignment = Alignment(horizontal="left", vertical="center")
                        row_idx += 1
                else:
                    row_data = [roster_date_str, cab_name, "-", "-", "No Booking"]
                    ws.append(row_data)
                    ws.row_dimensions[row_idx].height = 20
                    current_fill = zebra_fill if (row_idx % 2 == 0) else white_fill
                    for col_idx in range(1, 6):
                        c = ws.cell(row=row_idx, column=col_idx)
                        c.fill = current_fill
                        c.font = Font(name="Calibri", size=11)
                        c.border = thin_border
                        c.alignment = Alignment(horizontal="center", vertical="center")
                    row_idx += 1

            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    val_str = str(cell.value or '')
                    if len(val_str) > max_len:
                        max_len = len(val_str)
                ws.column_dimensions[col_letter].width = max(max_len + 5, 15)

            wb.save(file_name)
            excel_success = True
        except Exception as e:
            excel_success = False
            st.warning(f"Note: Excel file couldn't be updated (Is it open in Excel?): {e}")

        st.success("✅ Roster generated successfully!")

        # --- Display Results ---
        st.markdown("### 📋 Copy WhatsApp Message Below:")
        st.code(wa_output, language="text")

        # Download buttons
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            with open(txt_filename, "rb") as f:
                st.download_button(
                    label="📥 Download Notepad (.txt)",
                    data=f,
                    file_name=txt_filename,
                    mime="text/plain"
                )
        with col_d2:
            if os.path.exists(file_name):
                with open(file_name, "rb") as f:
                    st.download_button(
                        label="📊 Download Master Excel File",
                        data=f,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
