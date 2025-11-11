# Distribution Guide: Sharing PDF Reports Without ReportLab Installation

**Problem**: Other users want PDF output but don't want to install ReportLab

**Solutions**: 3 options, ranked by ease of use

---

## ⭐ RECOMMENDED: Option 1 - Portable Script (Easiest)

**File**: [geotable_pdf_generator_portable.py](geotable_pdf_generator_portable.py)

### How It Works

The script automatically detects if ReportLab is installed:
- **With ReportLab** → Generates PDF
- **Without ReportLab** → Generates formatted TEXT

### For You (Script Distributor)

**Just share one file:**
```
geotable_pdf_generator_portable.py
```

That's it! Users don't need to do anything special.

### For Users

**In Dynamo:**
1. Add Python Script node
2. Load [geotable_pdf_generator_portable.py](geotable_pdf_generator_portable.py)
3. Connect inputs (same as always):
   - `IN[0]` → Geotable data
   - `IN[1]` → Output path (e.g., `"C:/Reports/report.pdf"`)
   - `IN[2]` → `"auto"`
4. Run

**Output:**
- If they have ReportLab: Gets PDF file
- If they don't: Gets TEXT file (with note in output)

### Advantages
✅ No installation required
✅ Works immediately
✅ Falls back gracefully
✅ Single file to distribute

### Disadvantages
❌ Text format not as polished as PDF
❌ Users without ReportLab get different output

---

## Option 2 - One-Time Installer Script

**Files**:
- [install_reportlab_dynamo.py](install_reportlab_dynamo.py) (installer)
- [geotable_pdf_generator.py](geotable_pdf_generator.py) (main script)

### For You (Script Distributor)

**Share two files:**
1. `install_reportlab_dynamo.py` (installer)
2. `geotable_pdf_generator.py` (PDF generator)

**Provide these instructions:**
> "First time setup: Run install_reportlab_dynamo.py once, then use geotable_pdf_generator.py"

### For Users

**One-time setup (takes 30 seconds):**
1. Add Python Script node
2. Load [install_reportlab_dynamo.py](install_reportlab_dynamo.py)
3. Run (no inputs needed)
4. Wait for "✓ ReportLab installed successfully!"
5. Delete that node

**Normal use:**
1. Add Python Script node
2. Load [geotable_pdf_generator.py](geotable_pdf_generator.py)
3. Connect inputs
4. Run → Get PDF

### Advantages
✅ Automated installation
✅ One-time setup
✅ Professional PDF output for all users
✅ No manual command line work

### Disadvantages
❌ Requires initial setup step
❌ May fail if user doesn't have admin rights

---

## Option 3 - Bundled Package (Most Professional)

Create a complete package with ReportLab included.

### Setup (One-Time, by You)

**Step 1: Create package folder structure**
```
PDFReportGenerator/
├── bin/
│   └── geotable_pdf_generator.py
├── lib/
│   └── (ReportLab installed here)
└── README.txt
```

**Step 2: Install ReportLab to package folder**
```bash
pip install --target="PDFReportGenerator/lib" reportlab
```

**Step 3: Modify script to use bundled ReportLab**

Add this to the top of `geotable_pdf_generator.py`:
```python
import sys
import os

# Add bundled lib to path
script_dir = os.path.dirname(os.path.abspath(__file__))
lib_path = os.path.join(script_dir, '..', 'lib')
if os.path.exists(lib_path):
    sys.path.insert(0, lib_path)
```

**Step 4: Zip the entire folder**
```
PDFReportGenerator.zip
```

### For You (Script Distributor)

**Share one zip file:**
```
PDFReportGenerator.zip (contains script + ReportLab)
```

**Instructions for users:**
> "Extract the zip to any folder, then use bin/geotable_pdf_generator.py in Dynamo"

### For Users

**Setup:**
1. Extract `PDFReportGenerator.zip` to any location
2. Remember the path

**In Dynamo:**
1. Add Python Script node
2. Browse to extracted folder: `PDFReportGenerator/bin/geotable_pdf_generator.py`
3. Connect inputs
4. Run → Get PDF

### Advantages
✅ Zero installation
✅ Guaranteed to work
✅ Professional PDF output
✅ No dependencies
✅ Can be shared via network drive

### Disadvantages
❌ Larger file size (~5-10 MB)
❌ Requires initial setup by you
❌ Users need to extract zip

---

## Comparison Table

| Feature | Portable Script | Installer Script | Bundled Package |
|---------|----------------|------------------|-----------------|
| **File Size** | 50 KB | 50 KB + 5 KB | ~10 MB |
| **Setup Time (User)** | 0 seconds | 30 seconds | 1 minute |
| **PDF Output** | If ReportLab installed | After one-time install | Always |
| **Installation Needed** | No | Yes (automated) | No |
| **Admin Rights Needed** | No | Sometimes | No |
| **Network Drive Compatible** | Yes | No (install to local) | Yes |
| **Ease of Distribution** | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐ | ⭐⭐⭐ |
| **Ease of Use (User)** | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐ | ⭐⭐⭐⭐ |

---

## Recommended Approach

### For Small Teams (< 10 people)
**Use: Option 1 - Portable Script**

Why:
- Simplest distribution (1 file)
- Works immediately
- If someone wants PDF, they can install ReportLab themselves
- Most users are okay with text format for QC

### For Larger Teams (10-50 people)
**Use: Option 2 - Installer Script**

Why:
- Professional PDF output for everyone
- Automated installation (no manual steps)
- Only takes 30 seconds per user
- Centralized control

### For Enterprise / Network Drive Deployment
**Use: Option 3 - Bundled Package**

Why:
- No installation at all
- Deploy to shared network location
- Everyone uses same path
- IT department friendly
- Guaranteed consistent environment

---

## Example Deployment Scenarios

### Scenario 1: Email Distribution
**Best: Portable Script**

Email message:
```
Subject: New PDF Report Generator

Hi team,

Attached is the new report generator script.

Usage:
1. Save attached geotable_pdf_generator_portable.py to your computer
2. In Dynamo, add Python Script node and load it
3. Connect your geotable data and output path
4. Run!

Note: If you want PDF output instead of text, install ReportLab once:
  pip install reportlab

Questions? Let me know!
```

### Scenario 2: Network Drive Deployment
**Best: Bundled Package**

Setup:
```
\\CompanyShare\DynamoTools\PDFReportGenerator\
    bin\geotable_pdf_generator.py
    lib\reportlab\...
    README.txt
```

Email message:
```
Subject: PDF Report Generator Now Available

Hi team,

The PDF report generator is now available on the network drive.

Location: \\CompanyShare\DynamoTools\PDFReportGenerator\bin\geotable_pdf_generator.py

Usage:
1. In Dynamo, add Python Script node
2. Browse to the network path above
3. Connect inputs and run

No installation needed! See README.txt for details.
```

### Scenario 3: IT-Managed Rollout
**Best: Installer Script + Package Manager**

Have IT create a Dynamo package:
```
Company.PDFReports\
    bin\install.py (one-time)
    bin\geotable_pdf_generator.py (main script)
    docs\README.md
```

Deployed via Dynamo Package Manager - appears in everyone's Dynamo automatically.

---

## Testing Before Distribution

Whichever option you choose, test it first:

**Test 1: Clean Environment**
- Use a VM or test machine without ReportLab
- Verify your solution works as expected

**Test 2: User Simulation**
- Have a colleague try your instructions
- Watch for confusion or errors
- Refine instructions

**Test 3: Network Path**
- If using network drive, test with UNC paths
- Verify performance over network

---

## Support Documentation Template

**File to include: USAGE.txt**

```
PDF REPORT GENERATOR FOR CIVIL 3D DYNAMO
=========================================

QUICK START
-----------
1. Open your Dynamo graph
2. Add Python Script node
3. Load: geotable_pdf_generator_portable.py
4. Connect:
   IN[0] → Your geotable data
   IN[1] → Output path (e.g., "C:/Reports/report.pdf")
   IN[2] → "auto"
5. Run!

OUTPUT
------
- If you have ReportLab: PDF file
- If you don't: Text file (same format)

WANT PDF OUTPUT?
----------------
Install ReportLab once (takes 30 seconds):

1. Open Command Prompt
2. Run: pip install reportlab
3. Done!

Or use our installer: install_reportlab_dynamo.py

TROUBLESHOOTING
---------------
Issue: "Error: ReportLab is required"
→ Use install_reportlab_dynamo.py OR
→ Script will generate text file instead

Issue: "Module not found"
→ Ensure script path is correct
→ Check file wasn't renamed

SUPPORT
-------
Contact: [Your email]
Documentation: [Link to docs]
```

---

## Summary

| **Your Situation** | **Use This** |
|-------------------|-------------|
| Quick sharing with 1-5 people | **Portable Script** (Option 1) |
| Department rollout (10-50 people) | **Installer Script** (Option 2) |
| Enterprise deployment (50+ people) | **Bundled Package** (Option 3) |
| Network drive deployment | **Bundled Package** (Option 3) |
| Email attachment | **Portable Script** (Option 1) |
| IT-managed distribution | **Bundled Package** (Option 3) or **Installer** (Option 2) |

**Start with Option 1 (Portable Script)** - it's the simplest and works for 90% of cases. Upgrade to Options 2 or 3 only if needed.

---

## Files Reference

| File | Purpose | Use When |
|------|---------|----------|
| [geotable_pdf_generator_portable.py](geotable_pdf_generator_portable.py) | Auto-fallback to text | Distributing to users who may not have ReportLab |
| [geotable_pdf_generator.py](geotable_pdf_generator.py) | Original PDF generator | All users will have ReportLab installed |
| [install_reportlab_dynamo.py](install_reportlab_dynamo.py) | One-time installer | Helping users install ReportLab easily |
| [geotable_report_formatter.py](geotable_report_formatter.py) | Text-only formatter | Don't need PDF at all, text is fine |

---

**Questions about distribution? Check the files above or refer to [PDF_SETUP_GUIDE.md](PDF_SETUP_GUIDE.md) for more details.**
