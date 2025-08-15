' ******************************************************************************
' Auto-Generate Scrub Daddy engineering documentation from artwork
' Harshil Patel
' July 9, 2025
' Scrub Daddy, Inc.
'
' Upon executing the script a UserForm is launched that gathers the necessary
' information for generating the appropriate CAD model and drawings
'
' All files are named based on specifications and saved per the UserForm
'
' ******************************************************************************

Option Explicit

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim configArray As Variant
Dim configCount As Long
Dim swConfig As SldWorks.Configuration
Dim currentConfig As String
Dim boolStatus As Boolean
Dim longErrors As Long, longWarnings As Long

Sub main()
    
    newProductSetup.Caption = "Specialty Shape Generator v 1.1"

    ' Show the UserForm that will gather all the necessary information to execute the subsequent script
    newProductSetup.Show

    ' Call the first generateCAD macro to import the DXF/DWG and create a CAD model of the new foam product
    Call generateCAD1.main
        
    ' Reveal the exection progress page while hiding others
    newProductSetup.MultiPage1.Pages(3).Visible = True
    
    ' Switch to the progress execution page
    newProductSetup.MultiPage1.Value = 3
        
    ' Bring the userForm back
    newProductSetup.Show vbModeless
    
End Sub
