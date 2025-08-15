' ******************************************************************************
' newProductSetup UserForm
' Harshil Patel
' July 10, 2025
' Scrub Daddy, Inc.
'
' Macros to initialize and control the functionality of elements on the UserForm
'
' ******************************************************************************


Private Sub executeResumeButton_Click()
    Call b_main.main
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Prevent closing if the close mode is by user action (e.g., clicking X)
    If CloseMode = vbFormControlMenu Then
        Cancel = True ' Cancel the close operation
        MsgBox "Use the 'Cancel' button to close this Macro."
    End If
End Sub

Private Sub generateCADMacroBtnA_Click()
    ' hide the userform temporarily
    newProductSetup.Hide
    
    ' check if artwork file has been selected, prompt if otherwise
    If artworkFilePath.Value = "" Then
        MsgBox ("You must select an artwork file")
        Exit Sub
    End If
    
    ' check if save directory has been selected, prompt if otherwise
    If saveDirectory = "" Then
        MsgBox ("You must select a save directory")
        Exit Sub
    End If
    
    ' If we have the artwork file and save directory, proceed with execution
    Call generateCAD1.main
    
'    ' Switch to the Advanced page
'    newProductSetup.MultiPage1.Value = 2
'
'    ' Hide the execution progress page
'    newProductSetup.MultiPage1.Pages(3).Visible = False
'
'    ' Set focus on the button to complete CAD generation
'    newProductSetup.generateCADMacroBtnB.SetFocus
    
    ' Bring the userForm back
    newProductSetup.Show vbModeless
End Sub

Private Sub generateCADMacroBtnB_Click()
    ' Hide the userform so we can finish executing the script
    newProductSetup.Hide
    
    ' Call the second generateCAD macro to save the model into the save directory with the proper name
    Call generateCAD2.main
    
    Unload Me
    End
End Sub

Private Sub generateConfigsMacroBtn_Click()
    ' Check if any colors have been selected
    ' Create colorList array based on check boxes from the second page of the UserForm
    Dim colorListArray() As String
    Dim ctrl As Control
    Dim I As Integer
    Dim arrayLength As Integer
    I = 0
    For Each ctrl In Me.Controls
        ReDim Preserve colorListArray(I)
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.Value = True Then
                colorListArray(I) = ctrl.Caption
                I = I + 1
            End If
        End If
    Next ctrl
    ReDim Preserve colorListArray(UBound(colorListArray))
    arrayLength = UBound(colorListArray) - LBound(colorListArray)
    'Debug.Print UBound(colorListArray) & " - " & LBound(colorListArray) & " = " & arrayLength
    If arrayLength = 0 Then
        MsgBox ("Please select at least one color")
        Exit Sub
    End If
        
    ' Check if product name has been entered
    If productName = "" Then
        MsgBox ("Please enter a product (specialty shape) name")
        Exit Sub
    End If
    
    ' Check if engineer has been selected
    If engineerList.Text = "" Then
        MsgBox ("Please select an engineer")
        Exit Sub
    End If
    
    Call generateConfigs.main
    Unload Me
    End
End Sub

Private Sub generateDwgsMacroBtn_Click()
    ' check if artwork file has been selected, prompt if otherwise
    If artworkFilePath.Value = "" Then
        Dim response
        response = MsgBox("Do you want to select an artwork file?", vbYesNo Or vbDefaultButton1 Or vbQuestion, "No Artwork File Entered")
        If response = vbYes Then
            Exit Sub
        End If
    End If
    
    Call generateDrawing.main
    Unload Me
    End
End Sub

Private Sub resetButton_Click()
    Call UserForm_Initialize
End Sub

Private Sub selectAllBtn_Click()
    ' Selects all check boxes
    CheckBox1.Value = True  ' Red
    CheckBox2.Value = True  ' Orange
    CheckBox3.Value = True  ' Yellow
    CheckBox4.Value = True  ' Green
    CheckBox5.Value = True  ' Blue
    CheckBox6.Value = True  ' Pink
    CheckBox7.Value = True  ' Purple
    CheckBox8.Value = True  ' White
    CheckBox9.Value = True  ' Grey
    CheckBox10.Value = True ' Key Lime
    CheckBox11.Value = True ' Red Velvet

End Sub

Private Sub clearBtn_Click()
    ' Selects all check boxes
    CheckBox1.Value = False  ' Red
    CheckBox2.Value = False  ' Orange
    CheckBox3.Value = False  ' Yellow
    CheckBox4.Value = False  ' Green
    CheckBox5.Value = False  ' Blue
    CheckBox6.Value = False  ' Pink
    CheckBox7.Value = False  ' Purple
    CheckBox8.Value = False  ' White
    CheckBox9.Value = False  ' Grey
    CheckBox10.Value = False ' Key Lime
    CheckBox11.Value = False ' Red Velvet

End Sub

Private Sub UserForm_Initialize()
    ' Start on page 1
    newProductSetup.MultiPage1.Value = 0
    ' Hide the execution progress page
    newProductSetup.MultiPage1.Pages(3).Visible = False
    ' Resets the UserForm
    scrubDaddyOption.Value = True
    productName = ""
    artworkFilePath.Value = ""
    saveDirectory = ""
    engineerList.Clear
    ' Set the values the combo box should be populated with
    engineerList.AddItem "J. Sobel"
    engineerList.AddItem "E. Bennis"
    engineerList.AddItem "B.Rispoli"
    engineerList.AddItem "K.Babinchak"
    engineerList.AddItem "H. Patel"
    engineerList.Style = fmStyleDropDownList
    engineerList.BoundColumn = 0
    CheckBox1.Value = True  ' Red
    CheckBox2.Value = True  ' Orange
    CheckBox3.Value = True  ' Yellow
    CheckBox4.Value = True  ' Green
    CheckBox5.Value = True  ' Blue
    CheckBox6.Value = True  ' Pink
    CheckBox7.Value = True  ' Purple
    CheckBox8.Value = True  ' White
    CheckBox9.Value = True  ' Grey
    CheckBox10.Value = False ' Key Lime
    CheckBox11.Value = False ' Red Velvet
    
End Sub

Private Sub artworkBrowse_Click()
    Dim swApp As SldWorks.SldWorks
    Dim filter As String
    Dim getFile As String
    Dim fileConfig As String
    Dim fileDispName As String
    Dim fileDisplayState As String
    Dim fileOptions As Long
    Set swApp = Application.SldWorks
    filter = "DXF/DWG Files (*.dxf; *.dwg)|*.dxf;*.dwg|All Files (*.*)|*.*||"
    getFile = swApp.GetOpenFileName2("Select DXF/DWG file to use as dieline", "", filter, fileOptions, fileConfig, fileDispName, fileDisplayState)
    artworkFilePath.Text = getFile
End Sub

Private Sub cancelBtn_Click()
    Call UserForm_Initialize
    Unload Me
    End
End Sub

Private Sub generateBtn_Click()
    ' Hide the form so we can start executing the script and display the generated CAD
    Me.Hide
    
    ' Check if product name has been entered
    If productName = "" Then
        MsgBox ("Please enter a product (specialty shape) name")
        Me.Show
        Exit Sub
    End If
    
    ' check if artwork file has been selected, prompt if otherwise
    If artworkFilePath.Value = "" Then
        MsgBox ("Please select an artwork file")
        Me.Show
        Exit Sub
    End If
    
    ' check if save directory has been selected, prompt if otherwise
    If saveDirectory = "" Then
        MsgBox ("Please select a save directory")
        Me.Show
        Exit Sub
    End If
    
    ' Check if engineer has been selected
    If engineerList.Text = "" Then
        MsgBox ("Please select an engineer")
        Me.Show
        Exit Sub
    End If
    
    ' Check if any colors have been selected
    ' Create colorList array based on check boxes from the second page of the UserForm
    Dim colorListArray() As String
    Dim ctrl As Control
    Dim I As Integer
    Dim arrayLength As Integer
    I = 0
    For Each ctrl In Me.Controls
        ReDim Preserve colorListArray(I)
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.Value = True Then
                colorListArray(I) = ctrl.Caption
                I = I + 1
            End If
        End If
    Next ctrl
    ReDim Preserve colorListArray(UBound(colorListArray))
    arrayLength = UBound(colorListArray) - LBound(colorListArray)
    'Debug.Print UBound(colorListArray) & " - " & LBound(colorListArray) & " = " & arrayLength
    If arrayLength = 0 Then
        MsgBox ("Please select at least one color")
        Me.Show
        Exit Sub
    End If
    
End Sub

Private Sub saveDirectoryBrowse_Click()
    Dim shellApp As Object
    Dim folder As Object
    Set shellApp = CreateObject("Shell.Application")
    Set folder = shellApp.BrowseForFolder(0, "Where should this product be saved?", 0)
    saveDirectory.Text = folder.self.Path
End Sub
