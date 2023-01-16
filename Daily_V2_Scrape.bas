Attribute VB_Name = "Module1"
Sub Daily_V2_Scrape()
Attribute Daily_V2_Scrape.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'
    ActiveSheet.Name = "All Data"
    Sheets.Add(After:=Sheets("All Data")).Name = "Info"
    Sheets.Add(After:=Sheets("All Data")).Name = "Drilling"
    Sheets.Add(After:=Sheets("All Data")).Name = "Hourly Distribution"
    Sheets.Add(After:=Sheets("All Data")).Name = "Water"
    Sheets.Add(After:=Sheets("All Data")).Name = "Products"
    Sheets.Add(After:=Sheets("All Data")).Name = "Equipment"
    Sheets.Add(After:=Sheets("All Data")).Name = "Bits"
    Sheets.Add(After:=Sheets("All Data")).Name = "Personel"
    
    
'
    'BEGIN INFO DATA FILL'
    Sheets("All Data").Select
    Range("A4:N4").Select
    Selection.Copy
    Sheets("Info").Activate
        Range("A1").Select
        ActiveSheet.Paste Link:=True
        Range("E1").Select
        Application.CutCopyMode = False
        Selection.Cut Destination:=Range("A2")
        Range("G1").Select
        Selection.Cut Destination:=Range("B2")
        Range("I1").Select
        Selection.Cut Destination:=Range("A3")
        Range("K1").Select
        Selection.Cut Destination:=Range("A4")
        Range("A4").Select
        Selection.Cut Destination:=Range("B3")

    
    'Begin Drilling data fill'

    Sheets("All Data").Select
    Range("A7").Select
    Selection.Copy
    Range("A8").Select
    ActiveSheet.Paste Link:=True
    Range("B6:D6").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B8").Select
    ActiveSheet.Paste Link:=True
    Range("E8").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Drilling From"
    Range("F8").Select
    ActiveCell.FormulaR1C1 = "Drilling To"
    Range("G8").Select
    ActiveCell.FormulaR1C1 = "Reaming From"
    Range("H8").Select
    ActiveCell.FormulaR1C1 = "Reaming To"
    Range("K8:L8").Select
    ActiveCell.FormulaR1C1 = "Casing Size"
    Range("M8:N8").Select
    ActiveCell.FormulaR1C1 = "Casing Footage"
    Range("A8:N14").Select
    Selection.Copy
    Sheets("Drilling").Select
        Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=True, Transpose:=False
    
        Sheets("All Data").Select
        Range("A15:F42,A43:F44").Select
        Range("A44").Activate
        Application.CutCopyMode = False
        Selection.Copy
        Sheets("Hourly Distribution").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=True, Transpose:=True
    
    
    'BEGIN PERSONELL DATA FILL'
    
    Sheets("All Data").Select
    Range("A49:N52").Select
    Selection.Copy
    Sheets("Personel").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("All Data").Select
    Range("A53:N56").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Personel").Select
    Range("O1:AB4").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'BEGIN EQUIPMENT DATA FILL'
    
    Sheets("All Data").Select
    Range("G28:N38").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Equipment").Select
    Range("A1").Select
    ActiveSheet.Paste
    
    'BEGIN BIT DATA FILL'
    Sheets("All Data").Select
    Range("H42:N48").Select
    Selection.Copy
    Sheets("Bits").Select
    Range("A1").Select
    ActiveSheet.Paste
    
    'BEGIN HOULRY dISTRIBUTION FILL'
    
    
    Sheets("All Data").Select
    Range("A15:F44").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Hourly Distribution").Select
    Range("A1").Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=True
    
    'BEGIN PRODUCT DATA FILL'
    
    
    Sheets("All Data").Select
    Range("G16:N26").Select
    Selection.Copy
    Sheets("Products").Select
    Range("A1").Select
    ActiveSheet.Paste
    
    'BEGIN Water DATA FILL'
    Sheets("All Data").Select

    Range("K39:N39,K28:N28").Select
    Range("K28").Activate
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Water").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Select
    Application.CutCopyMode = False

    Range("A1").Select
    Selection.ClearContents
    
End Sub
