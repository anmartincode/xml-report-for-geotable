"""
Find Dynamo Python installation path
Run this in Dynamo to locate Python for ReportLab installation
"""

import sys
import os

def find_python():
    """Find Python executable and site-packages"""
    info = []

    # Python executable (may be AutoCAD)
    info.append(f"sys.executable: {sys.executable}")
    info.append("")

    # Python version
    info.append(f"Python version: {sys.version}")
    info.append("")

    # Search sys.path for Python installation
    info.append("Python paths (sys.path):")
    python_dirs = []
    for path in sys.path:
        info.append(f"  - {path}")

        # Look for Python3 or python.exe in path
        if 'Python3' in path or 'python' in path.lower():
            python_dirs.append(path)

            # Check if python.exe exists nearby
            parent_dir = os.path.dirname(path)
            python_exe = os.path.join(parent_dir, 'python.exe')
            if os.path.exists(python_exe):
                info.append(f"    → Found Python: {python_exe}")

    info.append("")

    # Try to find Dynamo installation
    info.append("Looking for Dynamo installation:")

    # Check common Dynamo paths
    possible_paths = [
        r"C:\Program Files\Dynamo\Dynamo Core\2.0\Python3\python.exe",
        r"C:\Program Files\Dynamo\Dynamo Core\2.10\Python3\python.exe",
        r"C:\Program Files\Dynamo\Dynamo Core\2.12\Python3\python.exe",
        r"C:\Program Files\Dynamo\Dynamo Core\2.13\Python3\python.exe",
        r"C:\Program Files\Dynamo\Dynamo Core\2.16\Python3\python.exe",
        r"C:\Program Files\Dynamo\Dynamo Core\2.17\Python3\python.exe",
        r"C:\Program Files\Dynamo\Dynamo Core\2.18\Python3\python.exe",
        r"C:\Program Files\Dynamo\Dynamo Core\3.0\Python3\python.exe",
    ]

    found_paths = []
    for path in possible_paths:
        if os.path.exists(path):
            found_paths.append(path)
            info.append(f"  ✓ {path}")

    if not found_paths:
        info.append("  ✗ No Dynamo Python found in standard locations")

    info.append("")

    # Installation command recommendation
    info.append("=" * 60)
    info.append("TO INSTALL REPORTLAB:")
    info.append("=" * 60)

    if found_paths:
        info.append("")
        info.append("Run this command in Command Prompt (as Administrator):")
        info.append("")
        for path in found_paths:
            info.append(f'"{path}" -m pip install reportlab')
            info.append("")
    else:
        info.append("")
        info.append("Option 1: Install to system Python")
        info.append("  pip install reportlab")
        info.append("")
        info.append("Option 2: Use portable script (no installation needed)")
        info.append("  Use: geotable_pdf_generator_portable.py")

    return '\n'.join(info)

# Dynamo execution
if 'IN' in dir():
    OUT = find_python()
else:
    OUT = "Run this script in Dynamo to find Python path"
