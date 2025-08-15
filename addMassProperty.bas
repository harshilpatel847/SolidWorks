' ******************************************************************************
' Update default unit system for the part
' Harshil Patel
' July 1, 2025
' Scrub Daddy, Inc.
'
' This macro sets the default unit system for the currently active part to be MMGS
' The number of significant figures for the mass properties is set to 6 so it's
' more appropriate for small parts with low densities
' (when using grams per millimeter cubed)
' ******************************************************************************

' Initiazation section where all variables and types are defined.
Dim swApp As Object
Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

Sub main()
    ' Sets the target in the DOM so the macro acts on the currently active part.
    Set swApp = Application.SldWorks
    Set Part = swApp.ActiveDoc
        
    ' Sets the unit system
    boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, 0, swUnitSystem_e.swUnitSystem_MMGS)
    ' Sets the number of significant figures for mass properties to 6
    boolstatus = Part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsMassPropDecimalPlaces, 0, 6)
    ' Rebuilds the part so the model properties recalculate based on the above changes.
    boolstatus = Part.EditRebuild3()
    
    ' Save the part.
    Dim swErrors As Long
    Dim swWarnings As Long
    boolstatus = Part.Save3(1, swErrors, swWarnings)

End Sub
