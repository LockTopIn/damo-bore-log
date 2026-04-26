
import streamlit as st
from datetime import datetime
import random
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import re

st.set_page_config(page_title="DAMO Bore Log Tool", layout="wide")

st.title("DAMO Bore Log Generator")

input_text = st.text_area("Paste Bore Input", height=300)


def parse_input(text):
    bores = []
    blocks = re.split(r'\n\s*\n', text.strip())
    for block in blocks:
        lines = [l.strip() for l in block.split("\n") if l.strip()]
        bore = {}
        m = re.match(r'(BR|PL)(\d+)\s*=\s*(\d+)', lines[0])
        if not m:
            continue
        bore["type"] = m.group(1)
        bore["name"] = f"{m.group(1)}{m.group(2)}"
        bore["footage"] = int(m.group(3))
        bore["lc"] = {}
        bore["eop_range"] = None  # FIX: default to None so PL without EOP won't crash

        for line in lines[1:]:
            if line.startswith("LC"):
                parts = line.replace("LC", "").split("=")
                rods = [int(x.strip()) for x in parts[0].split(",")]
                names = [x.strip() for x in parts[1].split(",")]
                for r, n in zip(rods, names):
                    bore["lc"][r] = n

            elif line.startswith("EOP"):
                eop_match = re.search(r'(\d+)-(\d+)', line)
                if eop_match:
                    bore["eop_range"] = (int(eop_match.group(1)), int(eop_match.group(2)))

            elif line.startswith("Depth"):
                nums = list(map(int, re.findall(r'\d+', line)))
                if bore["type"] == "BR":
                    bore["depth_range"] = (nums[0] * 12 + nums[1], nums[2] * 12 + nums[3])
                else:
                    # FIX: PL depth — handle "3'0"" format (2 numbers) or single number
                    if len(nums) >= 2:
                        bore["depth_flat"] = nums[0] * 12 + nums[1]
                    else:
                        bore["depth_flat"] = nums[0] * 12

        bores.append(bore)
    return bores


def inches_to_ft_in(val):
    return val // 12, val % 12


def rods_from_footage(ft):
    return ft // 10 + (1 if ft % 10 else 0)


def generate_eop(rods, lo, hi):
    eop = {}
    val = random.randint(lo, hi)
    for r in range(5, rods + 1, 5):
        val = max(lo, min(hi, val + random.choice([-1, 0, 1])))
        eop[r] = val
    return eop


def generate_br_depths(rods, low, high, lc):
    depths = []
    current = random.randint(36, 39)
    while len(depths) < rods:
        for _ in range(random.choice([3, 4, 5])):
            if len(depths) >= rods:
                break
            current += random.choice([-3, 3])
            current = max(low, min(high, current))
            depths.append(current)
        for _ in range(random.choice([2, 3])):
            if len(depths) >= rods:
                break
            depths.append(current)
    for r in lc:
        if r - 1 < len(depths):
            depths[r - 1] = random.randint(36, 42)
    return depths[:rods]


def generate_pl_depths(rods, flat):
    return[flat] * rods
    

def validate_depths(depths, bore_name):
    violations = []
    last_triple_end = -99  # tracks last 3-in-a-row ended

    for i in range(2, len(depths)):
        # check if 3 in a row
        if depths[i] == depths[i-1] == depths[i-2]:
            # checks if a 4th in a row (violation)
            if i >=3 and depths[i] == depths[i-3]:
                violations.append(f"Rod {i+1}: 4+ repeats of {depths[i]} inches")
            # checks cooldown - was there another triple too recently?
            if i - last_triple_end < 6:
                violations.append(f"ROD {i+1}: triple repeat too soon after last one")
            last_triple_end = i

    if violations:
        return f"{bore_name}: ❌ {len(violations)} violation(s) — " + " | ".join(violations)
    else:
        return f"{bore_name}: ✅ all rules pass"

def build_excel(bores):
    wb = Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
    title_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    number_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    thin = Side(style='thin')
    box_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for bore in bores:
        rods = rods_from_footage(bore["footage"])

        if bore["type"] == "BR":
            depths = generate_br_depths(rods, *bore["depth_range"], bore["lc"])
            result = validate_depths(depths, bore["name"])
            st.write(result)
        else:
            depths = generate_pl_depths(rods, bore["depth_flat"])

        # FIX: only generate EOP if range was provided
        if bore["eop_range"]:
            eop = generate_eop(rods, *bore["eop_range"])
        else:
            eop = {}

        for page_start in range(0, rods, 74):
            ws = wb.create_sheet(f"{bore['name']}_{page_start // 74 + 1}")

            # FIX: clean column dimension references (removed invisible unicode chars)
            ws.column_dimensions['A'].width = 6
            ws.column_dimensions['B'].width = 28
            ws.column_dimensions['C'].width = 6
            ws.column_dimensions['D'].width = 8
            ws.column_dimensions['E'].width = 8
            ws.column_dimensions['F'].width = 6
            ws.column_dimensions['G'].width = 28
            ws.column_dimensions['H'].width = 6
            ws.column_dimensions['I'].width = 8
            ws.column_dimensions['J'].width = 8

            ws.merge_cells("A1:J1")
            ws["A1"] = "DAMO Bore Log"
            ws["A1"].font = Font(bold=True, size=14)
            ws["A1"].fill = title_fill
            ws["A1"].alignment = Alignment(horizontal="center")

            headers = ["#", "Location Description", "EOP", "Depth (ft)", "Depth (in)"]

            for col in range(5):
                # LEFT SIDE (A–E)
                cell_left = ws.cell(row=3, column=col + 1, value=headers[col])
                cell_left.fill = number_fill if col == 0 else header_fill
                cell_left.font = Font(bold=True)
                cell_left.alignment = Alignment(horizontal="center")
                cell_left.border = box_border

                # RIGHT SIDE (F–J)
                cell_right = ws.cell(row=3, column=col + 6, value=headers[col])
                cell_right.fill = number_fill if col == 0 else header_fill
                cell_right.font = Font(bold=True)
                cell_right.alignment = Alignment(horizontal="center")
                cell_right.border = box_border

            for i in range(37):          # <- loop starts here
                left = page_start + i
                right = page_start + i + 37
                row_excel = 4 + i

                # LEFT SIDE
                for col in range(1, 11):
                    ws.cell(row=row_excel, column=col).border = box_border
                    
                if left < rods:
                    rod = left + 1
                    ft, inch = inches_to_ft_in(depths[left])
                    c = ws.cell(row=row_excel, column=1, value=rod)
                    c.fill = number_fill
                    c.border = box_border
                    ws.cell(row=row_excel, column=2, value=bore["lc"].get(rod, ""))
                    ws.cell(row=row_excel, column=3, value=eop.get(rod, ""))
                    ws.cell(row=row_excel, column=4, value=ft)
                    ws.cell(row=row_excel, column=5, value=inch)

                # RIGHT SIDE
                if right < rods:
                    rod = right + 1
                    ft, inch = inches_to_ft_in(depths[right])
                    c = ws.cell(row=row_excel, column=6, value=rod)
                    c.fill = number_fill
                    c.border = box_border
                    ws.cell(row=row_excel, column=7, value=bore["lc"].get(rod, ""))
                    ws.cell(row=row_excel, column=8, value=eop.get(rod, ""))
                    ws.cell(row=row_excel, column=9, value=ft)
                    ws.cell(row=row_excel, column=10, value=inch)

            # Total line on last page
            if page_start + 74 >= rods:
                ws.cell(row=42, column=9, value=f"Total: {bore['footage']}'")

    filename = f"DAMO_OUTPUT_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(filename)
    return filename


if st.button("Generate Bore Log"):
    if input_text.strip():
        try:
            bores = parse_input(input_text)
            if not bores:
                st.error("No valid bores found. Check your input format.")
            else:
                file = build_excel(bores)
                with open(file, "rb") as f:
                    st.download_button("Download Excel", f, file_name=file)
                st.success(f"Generated {len(bores)} bore(s) successfully!")
        except Exception as e:
            st.error(f"Error generating bore log: {e}")
    else:
        st.warning("Paste input first.")
