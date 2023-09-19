from fpdf import FPDF
import re
from collections import defaultdict

# Initialize PDF
pdf = FPDF()
pdf.add_page()
pdf.set_font("Arial", size=12)

# Extract entities and build hierarchy
with open('frmRTMAIN.vb', 'r', encoding='utf-8') as f:
    lines = f.readlines()

entity_hierarchy = defaultdict(list)
current_class = None
current_func = None

for line in lines:
    if 'Class ' in line:
        current_class = re.search(r'Class ([a-zA-Z0-9_]+)', line).group(1)
        pdf.cell(200, 10, f'Class: {current_class}', ln=True)
    elif 'Function ' in line or 'Sub ' in line:
        current_func = re.search(r'(Function|Sub) ([a-zA-Z0-9_]+)', line).group(2)
        entity_hierarchy[current_class].append(current_func)
        pdf.cell(200, 10, f'    Function/Sub: {current_func}', ln=True)
    elif 'Dim ' in line or 'Public ' in line or 'Private ' in line:
        var_name = re.search(r'(Dim|Public|Private) ([a-zA-Z0-9_]+)', line).group(2)
        pdf.cell(200, 10, f'        Variable: {var_name}', ln=True)
    elif "' " in line:
        comment = line.strip()
        pdf.cell(200, 10, f'        Comment: {comment}', ln=True)

# Generate PDF (Optional: Add more details, styling, etc.)
pdf.output('entity_hierarchy_extended.pdf')

# The PDF, entity_hierarchy_extended.pdf, will be generated in the current directory.
