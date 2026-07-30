Option Explicit

' UI helper macros for Elevated Water Tank workbook

Sub ClearInputs()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("01_Inputs")
    Application.ScreenUpdating = False
    ws.Unprotect Password:=ThisWorkbook.Worksheets("00_Info").Range("Macro_Password").Value
    On Error Resume Next
    Range("Tank_D").ClearContents
    Range("Tank_L").ClearContents
    Range("Tank_B").ClearContents
    Range("Water_Depth_H").ClearContents
    Range("Stage_Height_Hs").ClearContents
    Range("Concrete_Grade").ClearContents
    Range("Steel_Grade").ClearContents
    Range("Cover_nom").ClearContents
    Range("BarDia_Main_mm").ClearContents
    Range("BarDia_Dist_mm").ClearContents
    Range("User_Footing_Depth_mm").ClearContents
    Application.Calculate
    ws.Protect Password:=ThisWorkbook.Worksheets("00_Info").Range("Macro_Password").Value
    Application.ScreenUpdating = True
End Sub

Sub ExportOutputPDF()
    Dim ws As Worksheet
    Dim fname As String
    Set ws = ThisWorkbook.Worksheets("06_Output")
    fname = ThisWorkbook.Path & "\" & "Elevated_Water_Tank_Design_Keshri_Output.pdf"
    On Error GoTo ErrHandler
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fname, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    MsgBox "PDF exported to: " & fname, vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Error exporting PDF: " & Err.Description, vbExclamation
End Sub

Sub ProtectToggle()
    Dim ws As Worksheet
    Dim pwd As String
    pwd = ThisWorkbook.Worksheets("00_Info").Range("Macro_Password").Value
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "00_Info" Then
            If ws.ProtectContents Then
                ws.Unprotect Password:=pwd
            Else
                ws.Protect Password:=pwd
            End If
        End If
    Next ws
    MsgBox "Protection toggled (password from 00_Info).", vbInformation
End Sub

Sub InsertExample(caseType As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("01_Inputs")
    ws.Unprotect Password:=ThisWorkbook.Worksheets("00_Info").Range("Macro_Password").Value
    If LCase(caseType) = "circular" Then
        Range("Tank_Type").Value = "Circular"
        Range("Tank_D").Value = 20000
        Range("Water_Depth_H").Value = 6000
        Range("Stage_Height_Hs").Value = 12000
    ElseIf LCase(caseType) = "rectangular" Then
        Range("Tank_Type").Value = "Rectangular"
        Range("Tank_L").Value = 12000
        Range("Tank_B").Value = 8000
        Range("Water_Depth_H").Value = 4000
        Range("Stage_Height_Hs").Value = 8000
    End If
    Range("Concrete_Grade").Value = "M25"
    Range("Steel_Grade").Value = "Fe500"
    Range("Cover_nom").Value = 40
    Range("BarDia_Main_mm").Value = 16
    Range("BarDia_Dist_mm").Value = 8
    Application.Calculate
    ws.Protect Password:=ThisWorkbook.Worksheets("00_Info").Range("Macro_Password").Value
    MsgBox "Example data inserted: " & caseType, vbInformation
End Sub
