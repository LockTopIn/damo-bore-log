
import streamlit as st
from datetime import datetime
import random
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
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
        bore["type"] = m.group(1)
        bore["name"] = f"{m.group(1)}{m.group(2)}"
        bore["footage"] = int(m.group(3))
        bore["lc"] = {}

        for line in lines[1:]:
            if line.startswith("LC"):
                parts = line.replace("LC","").split("=")
                rods = [int(x.strip()) for x in parts[0].split(",")]
                names = [x.strip() for x in parts[1].split(",")]
                for r, n in zip(rods, names):
                    bore["lc"][r] = n

            elif line.startswith("EOP"):
                m = re.search(r'(\d+)-(\d+)', line)
                bore["eop_range"] = (int(m.group(1)), int(m.group(2)))

            elif line.startswith("Depth"):
                nums = list(map(int, re.findall(r'\d+', line)))
                if bore["type"] == "BR":
                    bore["depth_range"] = (nums[0]*12 + nums[1], nums[2]*12 + nums[3])
                else:
                    bore["depth_flat"] = nums[0]*12

        bores.append(bore)
    return bores

def inches_to_ft_in(val):
    return val//12, val%12

def rods_from_footage(ft):
    return ft//10 + (1 if ft%10 else 0)

def generate_eop(rods, lo, hi):
    eop = {}
    val = random.randint(lo, hi)
    for r in range(5, rods+1, 5):
        val = max(lo, min(hi, val + random.choice([-1,0,1])))
        eop[r] = val
    return eop

def generate_br_depths(rods, low, high, lc):
    depths = []
    current = random.randint(36,39)
    while len(depths) < rods:
        for _ in range(random.choice([3,4,5])):
            if len(depths) >= rods: break
            current += random.choice([-3,3])
            current = max(low, min(high, current))
            depths.append(current)
        for _ in range(random.choice([2,3])):
            if len(depths) >= rods: break
            depths.append(current)
    for r in lc:
        depths[r-1] = random.randint(36,42)
    return depths[:rods]

def generate_pl_depths(rods, flat):
    return [flat]*rods

def build_excel(bores):
    wb = Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
    title_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
	number_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6",fill_type="solid")
    
    for bore in bores:
        rods = rods_from_footage(bore["footage"])
        depths = generate_br_depths(rods, *bore["depth_range"], bore["lc"]) if bore["type"]=="BR" else generate_pl_depths(rods, bore["depth_flat"])
        eop = generate_eop(rods, *bore["eop_range"])

        for page_start in range(0, rods, 74):
            ws = wb.create_sheet(f"{bore['name']}_{page_start//74+1}")
            
            ws.column_dimensions‎[‎'A'].width = 6
            ws.column_dimensions['F'].width = 6

            ws.merge_cells("A1:J1")
            ws["A1"] = "DAMO Bore Log"
            ws["A1"].font = Font(bold=True, size=14)
            ws["A1"].fill = title_fill
            ws["A1"].alignment = Alignment(horizontal="center")

            headers = ["#","Location Description","EOP","Depth (ft)","Depth (in)"]
            
            for col in range(5):
                
                # LEFT SIDE (A-E)
                cell_left = ws.cell(row=3, column=col+1, value=headers[col‎])
                
                if col == 0:
                	cell_left.fill = number_fill
                else:
                    cell_left.fill = header_fill
                    
                # RIGHT SIDE (F-J)
                cell_right = ws.cell(row=3, column=col+6, value=headers[col])
                
                if col == 0:
                    cell_right.fill = number_fill
                else:
                    cell_right.fill = header_fill

            for i in range(37):
                left = page_start + i
                right = page_start + i + 37
                row_excel = 4 + i

                if left < rods:
                    rod = left+1
                    ft, inch = inches_to_ft_in(depths[left])
                    label = str(rod)
                    if rod in bore["lc"]:
                        label += "  " + bore["lc"][rod]

                	ws.cell(row=row_excel, column=1, value=rod).fill = number_fill
                	ws.cell(row=row_excel, column=2, value=bore‎["lc"‎].get(rod,""))
                	ws.cell(row=row_excel, column=3, value=eop.get(rod,""))
                	ws.cell(row=row_excel, column=4, value=ft)
                    ws.cell(row=row_excel, column=5, value=inch)

                if right < rods:
                    rod = right+1
                    ft, inch = inches_to_ft_in(depths[right])
                    label = str(rod)
                    if rod in bore["lc"]:
                        label += "  " + bore["lc"][rod]

					ws.cell(row=row_excel, column=6, value=rod).fill = number_fill
                    ws.cell(row=row_excel, column=7, value=bore‎["lc"].get(rod,""))
                    ws.cell(row=row_excel, column=8, value=eop.get(rod,""))
                    ws.cell(row=row_excel, column=9, value=ft)
                    ws.cell(row=row_excel, column=10, value=inch)

            if page_start + 74 >= rods:
                ws.cell(row=42, column=5, value=f"Total: {bore['footage']}'")

    filename = f"DAMO_OUTPUT_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(filename)
    return filename

if st.button("Generate Bore Log"):
    if input_text.strip():
        bores = parse_input(input_text)
        file = build_excel(bores)
        with open(file, "rb") as f:
            st.download_button("Download Excel", f, file_name=file)
    else:
        st.warning("Paste input first.")
