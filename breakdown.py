import os
import re
import shutil
from collections import defaultdict

def split_vb_file(file_path, output_folder):
    # Create output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)
    
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    block = []
    in_block = False
    all_imports = []
    block_specific_vars = defaultdict(list)
    block_name = ""
    
    # Collect all imports
    for line in lines:
        if 'Imports ' in line:
            all_imports.append(line.strip())
    
    # Parse lines to extract function and sub blocks
    for line in lines:
        if 'Function ' in line or 'Sub ' in line:
            in_block = True
            block_name = re.search(r'(Function|Sub) ([a-zA-Z0-9_]+)', line).group(2)
        if in_block:
            block.append(line.strip())
            
            # Collect variables specific to this block
            if 'Dim ' in line or 'Public ' in line or 'Private ' in line:
                block_specific_vars[block_name].append(line.strip())
                
        if 'End Function' in line or 'End Sub' in line:
            in_block = False
            
            # Write to a file
            output_path = os.path.join(output_folder, f"{block_name}.vb")
            with open(output_path, 'w', encoding='utf-8') as f_out:
                # Write comments for context
                f_out.write(f"' Extracted from {file_path}\n")
                
                # Write variables specific to this block
                for var in block_specific_vars[block_name]:
                    f_out.write(var + '\n')
                
                # Write the block content
                for b in block:
                    f_out.write(b + '\n')
            
            # Clear block for the next one
            block = []
    
    # Write all imports to a single file
    with open(os.path.join(output_folder, "AllImports.vb"), 'w', encoding='utf-8') as f:
        for imp in all_imports:
            f.write(imp + '\n')

# Define the input VB file path and the output folder where the split files will be saved
file_path = 'frmRTMAIN.vb'  # Replace with your VB file path
output_folder = 'frmRTMain_blocks'

# Perform the file splitting
split_vb_file(file_path, output_folder)

# Compress the generated files into a ZIP archive
#shutil.make_archive('output_blocks', 'zip', output_folder)
