# Sheet: 01_Inputs — Implementation Details

This file documents the exact Inputs worksheet layout, named ranges, data validations, and key cell formulas. Use this as a copy/paste blueprint when building the Excel sheet.

Worksheet name: 01_Inputs
Print area: A1:K60 (design UI tall but compact)

Top header
- A1: Project: (label)
- B1: Named range Project_Name (single cell)
- A2: Designed by: (label)  B2: "Keshri" (Named range Designed_By)

Section: Geometry (start row 4)
- A4: Tank Type (label)  B4: dropdown Data Validation list {"Circular","Rectangular"}
    - Named range: Tank_Type -> =01_Inputs!$B$4
- A5: Diameter D (mm)  B5: input (only used if Tank_Type="Circular")
    - Named range: Tank_D -> =01_Inputs!$B$5
- A6: Length L (mm)  B6: input (only used if Tank_Type="Rectangular")
    - Named range: Tank_L -> =01_Inputs!$B$6
- A7: Breadth B (mm)  B7: input (only used if Tank_Type="Rectangular")
    - Named range: Tank_B -> =01_Inputs!$B$7
- A8: Water depth H (mm)  B8: input -> Named range Water_Depth_H

Section: Staging (start row 10)
- A10: Stage type  B10: dropdown {"Column Staging","Shaft/Pedestal","Integral"} -> Named range Stage_Type
- A11: Stage height Hs (mm)  B11: input -> Named range Stage_Height_Hs

Section: Materials & Cover (start row 13)
- A13: Concrete grade  B13: dropdown (M20,M25,M30,M35,M40) -> Named range Concrete_Grade
- A14: Mapped fck (N/mm2) B14 formula: =VLOOKUP(Concrete_Grade, Tables!$A:$B,2,FALSE) -> Named range Concrete_fck
- A15: Steel grade  B15: dropdown (Fe415,Fe500,Fe550) -> Named range Steel_Grade
- A16: Mapped fy (N/mm2) B16 formula: =VLOOKUP(Steel_Grade, Tables!$D:$E,2,FALSE) -> Named range Steel_fy
- A17: Clear cover (mm) B17: input -> Named range Cover_nom

Section: Soil & Loads (start row 19)
- A19: Soil Type B19: dropdown {I,II,III} -> Named range Soil_Type
- A20: Allowable bearing (kN/m2) B20: input -> Named range Allowable_Bearing_kNpm2
- A21: Seismic Zone B21: dropdown {II,III,IV,V} -> Named range Seismic_Zone
- A22: Wind terrain category B22: dropdown {1,2,3} -> Named range Wind_Terrain

Section: Designer overrides & bars (start row 24)
- A24: User footing depth (mm) B24: input -> Named range User_Footing_Depth_mm
- A25: Main bar dia (mm) B25: dropdown (8,10,12,16,20,25,32) -> Named range BarDia_Main_mm
- A26: Distribution bar dia (mm) B26: dropdown (8,10,12,16) -> Named range BarDia_Dist_mm

Buttons (row 28)
- Insert Example Circular (Button) -> calls InsertExample("circular")
- Insert Example Rectangular -> InsertExample("rectangular")
- Clear Inputs -> ClearInputs()
- Export Output to PDF -> ExportOutputPDF()
- Toggle Protect -> ProtectToggle()

Interactive diagram
- Place a circle or rectangle shape on the right (E4:J18) and insert textboxes for D,L,B,H that reference named ranges (e.g. select a textbox and put formula =Tank_D). Format textboxes to show units (mm) and update automatically.

Conditional formatting & UI hints
- If Tank_Type="Circular" then show/enable Tank_D; hide/gray Tank_L and Tank_B. Use simple conditional formatting (font color gray) and data validation help text.
- Input cells unlocked and colored light green; calculation sheets locked/protected.

Validation rules (examples)
- Water_Depth_H > 0 and < 20000 mm (simple limit)
- Tank_D > 2000 and < 50000 mm
- Cover_nom between 20 and 100 mm

Named ranges summary (create via Name Manager):
- Project_Name, Designed_By, Tank_Type, Tank_D, Tank_L, Tank_B, Water_Depth_H, Stage_Type, Stage_Height_Hs, Concrete_Grade, Concrete_fck, Steel_Grade, Steel_fy, Cover_nom, Soil_Type, Allowable_Bearing_kNpm2, Seismic_Zone, Wind_Terrain, User_Footing_Depth_mm, BarDia_Main_mm, BarDia_Dist_mm, Design_Units

Notes
- Keep cell formulas minimal on the Inputs sheet — only value mapping lookups and unit conversion helpers. Heavy calculations live on 02_Loads and 03_Calculations.
- Protect sheet with password (Kesukriti@1998) when workbook is finalized.