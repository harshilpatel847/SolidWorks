' ******************************************************************************
' Check and Approve Macro
' Harshil Patel
' August 15, 2025
' Scrub Daddy, Inc.
'
' Add approving manager's name & date to drawing revision table and title block
'
' ******************************************************************************

Option Explicit
Dim swApp As Object
Dim Part As Object
Dim swDrawing As SldWorks.DrawingDoc
Dim swSheet As SldWorks.Sheet
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Dim chkName As SldWorks.note
Dim chkDate As SldWorks.note
Dim paperSize As swDwgPaperSizes_e
Dim width As Double
Dim height As Double
Dim note(2) As String
Dim swErrors As Long
Dim swWarnings As Long

Sub main()

    Set swApp = Application.SldWorks
    Set Part = swApp.ActiveDoc
    Set swDrawing = Part
    Dim myModelView As Object
    Set myModelView = Part.ActiveView
    Set swSheet = swDrawing.GetCurrentSheet
    paperSize = swSheet.GetSize(width, height)
    
    myModelView.FrameState = swWindowState_e.swWindowMaximized
    
    ' Check the sheet format size and update note IDs
    Select Case width
        Case 0.21  ' A4 sheet format
            note(0) = 248
            note(1) = 249
        Case 0.42  ' A3 sheet format
            note(0) = 350
            note(1) = 351
        Case 0.594 ' A2 sheet format
            note(0) = 374
            note(1) = 375
        Case 0.841 ' A1 sheet format
            note(0) = 575
            note(1) = 574
        Case Else  ' Not one of the standard formats
    End Select
    
'    If width > 0.4 And width < 0.5 Then ' A3 sheet format
'        note(0) = 0.2405
'        note(1) = 0.263
'        note(2) = 0.043
'    ElseIf width > 0.5 Then ' A2 sheet format
'        note(0) = 0.414
'        note(1) = 0.437
'        note(2) = 0.0475
'    ElseIf width < 0.4 Then ' A4 sheet format
'        note(0) = 0.031
'        note(1) = 0.053
'        note(2) = 0.047
'    End If
    
    ' Fill in the "Approved By/Date" field of the revision table
    boolstatus = Part.Extension.SelectByID2("", "REVISIONTABLE", width - 0.02, height - 0.02, 0, False, 0, Nothing, 0)
    Dim myTable As Object
    Set myTable = Part.SelectionManager.GetSelectedObject5(1)
    myTable.Text(2, 4) = "J. SOBEL/" & UCase(Format(Date, "ddmmmyyyy"))
    
    ' Edit the sheet format to update the title block fields
    boolstatus = Part.Extension.SelectByID2("Sheet1", "SHEET", 0, 0, 0, False, 0, Nothing, 0)
    Part.EditTemplate
    Part.EditSketch
    Part.ClearSelection2 True
    
    ' Fill in the checked by name field
    boolstatus = Part.Extension.SelectByID2("DetailItem" & note(0) & "@Sheet Format1", "NOTE", 0, 0, 0, False, 0, Nothing, 0)
    Set chkName = Part.SelectionManager.GetSelectedObject5(1)
    chkName.SetText "J. SOBEL"
    Part.ClearSelection2 True
    
    ' Fill in the checked by date field
    boolstatus = Part.Extension.SelectByID2("DetailItem" & note(1) & "@Sheet Format1", "NOTE", 0, 0, 0, False, 0, Nothing, 0)
    Set chkDate = Part.SelectionManager.GetSelectedObject5(1)
    chkDate.SetText UCase(Format(Date, "ddmmmyyyy"))
    Part.ClearSelection2 True
    
    Part.EditSheet
    Part.EditSketch

    ' Save the part.
    boolstatus = Part.Save3(1, swErrors, swWarnings)

End Sub
