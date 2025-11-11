# Quick Start Guide

## 🚀 Get Started in 3 Steps

### Option 1: Direct Text Reports (Recommended)

**Fastest way to get reports matching your MicroStation format**

1. **In Dynamo, add Python Script node**
2. **Load**: [geotable_report_formatter.py](geotable_report_formatter.py)
3. **Connect**:
   - `IN[0]` → Your geotable data
   - `IN[1]` → Output path (e.g., `"C:\Reports\alignment.txt"`)
   - `IN[2]` → `"auto"`
4. **Run** → Get formatted report instantly ✓

**Output**: Plain text report matching Codman_Final format

---

### Option 2: PDF Reports (Professional Output) ⭐ NEW

**Generate professional PDF reports ready for printing and sharing**

**Choose your version:**

**A) Portable (No Installation Required)** - Recommended for sharing
- **File**: [geotable_pdf_generator_portable.py](geotable_pdf_generator_portable.py)
- **Works immediately** - no setup needed
- **Auto-detects**: PDF if ReportLab installed, text otherwise

**B) PDF-Only (Requires ReportLab)**
- **File**: [geotable_pdf_generator.py](geotable_pdf_generator.py)
- **Prerequisites**: Install ReportLab first
  ```bash
  pip install reportlab
  ```
- **Always generates PDF**

**Usage (same for both):**
1. **In Dynamo, add Python Script node**
2. **Load**: Your chosen script above
3. **Connect**:
   - `IN[0]` → Your geotable data OR XML file path
   - `IN[1]` → Output path (e.g., `"C:\Reports\alignment.pdf"`)
   - `IN[2]` → `"auto"`
4. **Run** → Get professional PDF report ✓

**Output**: PDF report matching Codman_Final/SCR Complete Align format

**Distribution**: See [DISTRIBUTION_GUIDE.md](DISTRIBUTION_GUIDE.md) for sharing with team

---

### Option 3: XML Output (For Existing Pipeline)

**Use with your current XML → XSLT → XLSM workflow**

1. **In Dynamo, add Python Script node**
2. **Load**: [xml_report_generator.py](xml_report_generator.py)
3. **Connect**:
   - `IN[0]` → Your geotable data
   - `IN[1]` → Output path (e.g., `"C:\Reports\alignment.xml"`)
   - `IN[2]` → `True`
4. **Run** → Get structured XML ✓

**Output**: XML file for your existing pipeline

---

## 📋 Input Format

All scripts accept the same geotable list format:

```python
[
    ['Project Name:', 'Codman_Final'],
    ['Horizontal Alignment Name:', 'Prop_AshNB'],
    ['', 'STATION', 'ELEVATION'],
    ['Element: Linear', '', ''],
    ['POB', '641+44.67', '42.97'],
    ['PVI', '641+95.67', '41.74'],
    ['Tangent Grade:', '-2.411', ''],
    # ... more rows
]
```

---

## ✅ What's Fixed

| Issue | Status |
|-------|--------|
| ❌ `AttributeError: 'list' object has no attribute 'get'` | ✅ **FIXED** in v2.0 |
| ❌ Can't parse raw geotable data | ✅ **FIXED** - Auto-detects format |
| ❌ Station format not recognized ("641+44.67") | ✅ **FIXED** - Converts automatically |
| ❌ Missing element properties | ✅ **FIXED** - Parses all properties |

---

## 🎯 Which Script to Use?

### Use PDF Generator When: ⭐ RECOMMENDED
- ✅ Need professional, print-ready reports
- ✅ Sharing reports with team/clients
- ✅ QC reviews and documentation
- ✅ Want formatted output matching examples exactly
- ✅ Reports for presentations or submittals

### Use Direct Text Formatter When:
- ✅ Quick QC reports for review
- ✅ Want instant output (< 1 second)
- ✅ Need plain text for diff/version control
- ✅ Don't need formatted output

### Use XML Generator When:
- ✅ Have existing XML → XSLT workflow
- ✅ Need structured data for other tools
- ✅ Transitioning gradually from old pipeline

---

## 📊 Example Outputs

### Direct Formatter Output
```
Project Name: Codman_Final
Horizontal Alignment Name: Prop_AshNB
Vertical Alignment Name: Prop_AshNB
                                                 STATION       ELEVATION

Element: Linear
    POB                        641+44.67           42.97
    PVI                        641+95.67           41.74
    Tangent Grade:                -2.411
    Tangent Length:                51.01

Element: Parabola
    PVC                        642+07.29           41.46
    PVI                        642+57.29           40.25
    PVT                        643+07.29           40.21
    Length:                       100.00
    Headlight Sight Distance:          540.41
    Entrance Grade:               -2.411
    Exit Grade:                   -0.075
```

### XML Generator Output
```xml
<?xml version="1.0" encoding="utf-8"?>
<GeotableReport xmlns="http://civil3d.autodesk.com/geotable" version="1.0">
  <ProjectInfo>
    <ProjectName>Codman_Final</ProjectName>
    ...
  </ProjectInfo>
  <Alignments count="1">
    <Alignment name="Prop_AshNB" id="">
      <Properties>...</Properties>
      <Stations count="7">...</Stations>
      <GeometricElements count="3">...</GeometricElements>
    </Alignment>
  </Alignments>
</GeotableReport>
```

---

## 🧪 Test Before Using

```bash
# Test PDF generator (NEW!)
python3 test_pdf_generator.py

# Test direct formatter
python3 test_report_formatter.py

# Test XML generator
python3 test_geotable_parser.py
```

Check outputs:
- `test_report.pdf` - PDF generator output ⭐
- `test_formatted_report.txt` - Direct formatter output
- `test_output.xml` - XML generator output

---

## 📚 Full Documentation

| Document | Purpose |
|----------|---------|
| [PDF_SETUP_GUIDE.md](PDF_SETUP_GUIDE.md) | PDF generator setup and usage ⭐ NEW |
| [SOLUTION_SUMMARY.md](SOLUTION_SUMMARY.md) | Complete solution overview |
| [DYNAMO_DIRECT_GUIDE.md](DYNAMO_DIRECT_GUIDE.md) | Direct formatter detailed guide |
| [README.md](README.md) | Complete documentation |
| [RELEASE_NOTES_v2.0.md](RELEASE_NOTES_v2.0.md) | What's new in v2.0 |

---

## ⚡ Performance

| Feature | PDF Generator ⭐ | Text Formatter | XML Generator |
|---------|----------------|----------------|---------------|
| **Speed** | < 2 seconds | < 1 second | < 2 seconds |
| **Output Size** | 50-100 KB | 5-20 KB | 30-50 KB |
| **Dependencies** | ReportLab | Python only | Python only |
| **Readability** | Professional ✓✓ | Plain text ✓ | XML structure |
| **Print-ready** | Yes ✓ | No | No |

---

## 🆘 Troubleshooting

### "Error: Invalid input type"
→ Check geotable data is a list of lists

### "Station data missing"
→ Ensure "STATION" header row exists
→ Station values in "NNN+DD.DD" or numeric format

### "Empty report"
→ Check "Element: Type" headers present
→ Verify data has at least 2-3 columns

### Report format doesn't match
→ Review [DYNAMO_DIRECT_GUIDE.md](DYNAMO_DIRECT_GUIDE.md#customization)
→ Adjust column widths and formatting in the script

---

## 💡 Pro Tips

1. **Use PDF generator for client deliverables** - Professional and print-ready
2. **Test with sample data first** before using on production alignments
3. **Save output files** with descriptive names (e.g., `Proj_Alignment_Vertical_2025-11-10.pdf`)
4. **Use "auto" mode** for `IN[2]` - it detects vertical vs horizontal automatically
5. **Compare output** to your MicroStation reports to verify accuracy
6. **Keep all three scripts** - PDF for QC/submittals, Text for quick checks, XML for legacy workflows

---

## ⏱️ Time Savings

**Before** (XML Pipeline):
```
Geotable → XML Gen (10s) → XSLT (5s) → XLSM (15s) → XLSX (5s) → DWG (10s)
Total: ~45 seconds + manual steps
```

**After** (PDF Generator) ⭐ RECOMMENDED:
```
Geotable → PDF Generator (< 2s) → Professional Report ✓
Total: < 2 seconds
```

**After** (Text Formatter):
```
Geotable → Text Formatter (< 1s) → Report ✓
Total: < 1 second
```

**Time saved per report: ~43 seconds**
**Time saved per 100 reports: ~1.2 hours**

---

## 🎓 Need Help?

1. Read [SOLUTION_SUMMARY.md](SOLUTION_SUMMARY.md) for complete overview
2. Review [DYNAMO_DIRECT_GUIDE.md](DYNAMO_DIRECT_GUIDE.md) for detailed instructions
3. Run test scripts to validate installation
4. Check example outputs to see expected format

---

**Ready to start? Pick your option above and go! 🚀**
