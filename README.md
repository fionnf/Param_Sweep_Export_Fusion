# ParamSweepExport

This Fusion 360 script automates the creation and export of parameter-swept part variants.

## What it does
- Iterates over two user parameters:
  - clr → unitless ratio (e.g. 1.00, 1.20, 1.40 …)
  - dpt → dimension in millimetres (e.g. 0.5, 0.6 …)
- Assigns a letter code (A, B, C …) to each variant.
- Replaces any sketch text equal to `XXX` with the letter code, so each version is embossed with its identifier.
- Exports each version as an STL file named:
  COMPONENT_clr<value>_dpt<value>mm_<LetterCode>.stl
- Logs all results to a variants.csv file with columns:
  code, clr, dpt, filename
- Restores the original parameters and placeholder text when done.

## How to use
1. Open or create a Fusion 360 design (.f3d) and save it.
2. Define two User Parameters named `clr` (unitless) and `dpt` (length in mm).
3. Add a sketch text object with the text exactly set to `XXX`. This is the placeholder that will be replaced by the letter code.
4. Place the script folder `ParamSweepExport` into your Fusion 360 Scripts directory:
   - macOS: ~/Library/Application Support/Autodesk/Autodesk Fusion 360/API/Scripts/
   - Windows: %APPDATA%\Autodesk\Autodesk Fusion 360\API\Scripts\
5. Run the script from Tools → Scripts and Add-Ins… → My Scripts → ParamSweepExport → Run.
6. STL files and the variants.csv log will appear in the configured OUTPUT_DIR.

## Output
- One STL per parameter combination.
- variants.csv containing a record of parameter values and corresponding letter codes.
