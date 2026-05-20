# ppt2a4 – PowerPoint to A4 Sheet Generator

**Automatically export PowerPoint slides and arrange them on A4 sheets (6 or 9 slides per sheet) with full control over layout. No manual positioning in Photoshop.**

**Download ready-to-use .exe**
**Go to Releases and download PPT_to_A4_Generator.exe**
**Just run it – Microsoft PowerPoint must be installed.**

## Features
- One‑click generation – export slides, rotate, arrange, save as PNG + optional PDF
- Auto layout – 3 columns, right‑to‑left order, adjustable margins
- Smart orientation – landscape slides are rotated, portrait kept intact
- Slide numbering (optional) – global counter when merging multiple presentations
- Custom footer text – add any text at the bottom of each sheet
- Merge several PPTX files into one continuous sequence
- Preview first sheet before generating everything
- Flexible output naming – use {name} and {n} placeholders

## Requirements for the .exe
- Windows 7 / 10 / 11
- Microsoft PowerPoint (2007 or later)

## Build from source (optional)
**If you prefer to run the Python script or build your own .exe:**
1. Clone this repository
2. Install dependencies:
   pip install -r requirements.txt
3. Run the script:
   python ppt2a4.py
4. Or build a standalone .exe using build.bat (requires PyInstaller).

## How to use
1. Launch ppt2a4.exe (or run the Python script).
2. Click Add Files – select one or more .pptx presentations.
3. Adjust settings:
   - Slides per sheet: Auto / 6 / 9
   - Margin (pixels between slides)
   - Footer text, numbering, PDF creation, filename template.
4. Choose an output folder.
5. Click Generate Sheets – wait for completion.
6. Find PNG sheets (and a combined PDF if selected) in the output folder.

Example output
- my_presentation_sheet_1.png – first A4 sheet (slides 1–6 or 1–9)
- my_presentation_sheet_2.png – second sheet
- my_presentation_combined.pdf – all sheets in one PDF (optional)

License
MIT

Author
night-rise (https://github.com/night-rise)