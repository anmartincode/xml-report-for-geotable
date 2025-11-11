"""
One-time ReportLab installer for Dynamo
Run this script once in Dynamo to install ReportLab

Instructions:
1. Add Python Script node in Dynamo
2. Load this file
3. Run once (no inputs needed)
4. Delete the node after installation
"""

import sys
import subprocess

def install_reportlab():
    """Install ReportLab using pip"""
    try:
        # Get Python executable path
        python_exe = sys.executable

        # Try to install ReportLab
        result = subprocess.run(
            [python_exe, '-m', 'pip', 'install', 'reportlab'],
            capture_output=True,
            text=True
        )

        if result.returncode == 0:
            return f"✓ ReportLab installed successfully!\n\nPython path: {python_exe}\n\nYou can now use the PDF generator."
        else:
            return f"✗ Installation failed:\n{result.stderr}\n\nTry manual installation:\n{python_exe} -m pip install reportlab"

    except Exception as e:
        return f"✗ Error: {str(e)}\n\nManual installation:\n1. Open Command Prompt as Administrator\n2. Run:\n   \"{sys.executable}\" -m pip install reportlab"

# Dynamo execution
if 'IN' in dir():
    OUT = install_reportlab()
else:
    OUT = "Run this script in Dynamo to install ReportLab"
