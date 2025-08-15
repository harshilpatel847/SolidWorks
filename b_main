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

    ' Hide the userform so we can finish executing the script
    newProductSetup.Hide
    
    ' Call the second generateCAD macro to save the model into the save directory with the proper name
    Call generateCAD2.main
    
    ' Call the autoGenerateColors macro to create custom configurations for all available colors
    Call generateConfigs.main
    
    ' Create a drawing for each configuration in the model
    ' Call the generateProductDrawing macro for each configuration
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    configArray = swModel.GetConfigurationNames
    For configCount = 0 To UBound(configArray)
        currentConfig = configArray(configCount)
        boolStatus = swModel.ShowConfiguration2(currentConfig)
        If currentConfig <> "Default" Then
            Call generateDrawing.main
        End If
    Next configCount
    
    ' Save
    boolStatus = swModel.Save3(1, longErrors, longWarnings)
    
    ' Unload the UserForm from memory
    Unload newProductSetup
    
    ' Notify the user that macro execution is complete
    MsgBox "Execution complete"
    
End Sub

