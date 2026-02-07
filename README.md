# Comprehensive SOP Presentation Generator

This repository contains a Python script to generate a comprehensive Standard Operating Procedures (SOP) PowerPoint presentation from multiple DOCX source documents.

## Overview

The `generate_comprehensive_sop.py` script creates a professional PowerPoint presentation that:
- Uses the styling, fonts, and layout from `SOP Draft.pptx` as a template
- Creates detailed coverage (2+ slides) for major SOPs:
  - Employee Onboarding
  - Employee Relieving
  - Leave & Attendance Policy
  - Performance Appraisal
- Creates summary coverage (1 slide) for supporting SOPs:
  - Payroll & Benefits
  - Training & Development
  - Disciplinary & Grievance
  - NDA & Legal Agreements
  - Organization Structure

## Requirements

The script requires the following Python packages:
```bash
pip install python-pptx python-docx
```

## Usage

Simply run the script:
```bash
python3 generate_comprehensive_sop.py
```

The script will:
1. Load the `SOP Draft.pptx` template
2. Extract content from all SOP DOCX files
3. Generate a comprehensive presentation with ~23 slides
4. Save the output as `Comprehensive_SOP_Presentation.pptx`

## Output Structure

The generated presentation includes:
- 1 title slide
- 8 detailed slides (4 SOPs × 2 slides each)
- 5 summary slides (5 SOPs × 1 slide each)
- 9 original template slides

Total: ~23 slides maintaining consistent styling with the template

## File Structure

```
.
├── SOP Draft.pptx                           # Template presentation
├── generate_comprehensive_sop.py            # Generation script
├── Comprehensive_SOP_Presentation.pptx      # Generated output
├── P9_IND_HR_002-00_Employee_Onboarding_SOP(1).docx
├── P9_IND_HR_003-00_Employee_Relieving_SOP.docx
├── P9_IND_HR_004-00_Leave and attendance_Policy_SOP.docx
├── P9_IND_HR_005-00_Performance _Apraisal_SOP.docx
├── P9_IND_HR_006-00_Payroll_benefits_SOP.docx
├── P9_IND_HR_007-00_Training_Development_SOP.docx
├── P9_IND_HR_008-00_Disciplinary_Grieveance_SOP.docx
├── P9_IND_HR_009-00_NDA_Legal agreements_SOP.docx
└── P9_IND_HR_010-00_Organogram_SOP.docx
```

## Styling

The generated presentation maintains all styling from the template:
- Font: Aptos
- Title font size: 32pt (bold)
- Content font size: 14pt
- All logos and design elements from the template
- Consistent slide layouts

## Notes

- The script automatically extracts content from DOCX files using headings and paragraphs
- Long content items are automatically truncated to fit slide constraints
- Each slide is limited to ~10 bullet points for readability
- The template's original slides are preserved in the output
