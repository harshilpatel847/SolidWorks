' ******************************************************************************
' Auto-Generate CAD Drawings for Part
' Harshil Patel
' July 11, 2025
' Scrub Daddy, Inc.
'
' This macro acts on a saved part to generate a specialty foam shape drawing
' The generated drawing uses a template based on the part's sub-category
' (Either a Scrub Mommy product or Scrub Daddy product)
'
' ******************************************************************************

Option Explicit

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swDrawing As SldWorks.DrawingDoc
Dim swConfigMgr As SldWorks.ConfigurationManager
Dim swConfig As SldWorks.Configuration
Dim swModelDocExt As SldWorks.ModelDocExtension
Dim swCustomPropMgr1 As SldWorks.CustomPropertyManager
Dim swCustomPropMgr2 As SldWorks.CustomPropertyManager
Dim form As Object
Dim currentSheet As Object
Dim swModView As SldWorks.ModelView
Dim swFrontView As SldWorks.View
Dim swRightView As SldWorks.View
Dim swIsoView As SldWorks.View
Dim swSheet As SldWorks.Sheet
Dim boolStatus As Boolean
Dim longStatus As Long
Dim longErrors As Long, longWarnings As Long
Dim templatePath As String
Dim templateFile As String
Dim viewType As Long
Dim subCategory As String
Dim valueResolved As String
Dim drawnBy As String
Dim description As String
Dim color As String
Dim wasResovled As Boolean
Dim linkToProp As Boolean
Dim sheetWidth As Double
Dim sheetHeight As Double
Dim sheetProperties As Variant
Dim width As Double
Dim height As Double
Dim thickness As Double
Dim baseNamePath As String
Dim partName As String
Dim saveDirectory As String
Dim isoNote As SldWorks.Note
Dim isoNoteText As String
Dim isoAnnotation As SldWorks.Annotation
Dim isoNoteFormat As SldWorks.TextFormat
Dim matlNote1 As SldWorks.Note
Dim matlNote2 As SldWorks.Note
Dim matlNoteAnnotation As SldWorks.Annotation
Dim matlNoteFormat As SldWorks.TextFormat
Dim myRevisionTable As Object
Dim selectData As SldWorks.selectData
Dim outlineFront() As Double
Dim outlineRight() As Double
Dim outlineIso() As Double
Dim myDisplayDim As Object
Dim viewPosFront() As Double
Dim viewPosRight() As Double
Dim viewPosIso() As Double
Dim isoWidth As Double
Dim isoHeight As Double
Dim swSketchManager As SldWorks.SketchManager
Dim sheetScale As Integer
Dim artworkFileFullPath As String
Dim artworkFileRelativePath As String
Dim artworkFileName As String
Dim artworkFileNote As SldWorks.Note
Dim artworkFileAnnotation As SldWorks.Annotation
Dim artworkFileNoteFormat As SldWorks.TextFormat
Dim response
Dim selectionEnum As Integer


Sub main()

    ' Sets the appropriate targets in the D.O.M. so the macro works on the specific document you have open and active right now.
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    ' Error handling in the event this macro is launched from the special execution page of the UserForm
    If swModel Is Nothing Then
        MsgBox ("Please open a specialty shape CAD model and try again (generateDrawing)")
        Exit Sub
    End If
    
    Set swModelDocExt = swModel.Extension
    Set swConfig = swModel.GetActiveConfiguration
    Set swCustomPropMgr2 = swModelDocExt.CustomPropertyManager("")
    Set swCustomPropMgr1 = swConfig.CustomPropertyManager

    ' Information about current part to use for drawing generation and subsequent drawing file save
    baseNamePath = Right(swModel.GetPathName, Len(swModel.GetPathName) - InStrRev(swModel.GetPathName, "\"))
    partName = Left(baseNamePath, InStrRev(baseNamePath, ".") - 1)
    saveDirectory = Left(swModel.GetPathName, Len(swModel.GetPathName) - Len(baseNamePath))
    
    ' Get the product metadata from FILE properties
    longStatus = swCustomPropMgr2.Get6("Sub Category", False, subCategory, valueResolved, wasResovled, linkToProp)
    longStatus = swCustomPropMgr2.Get6("DrawnBy", False, drawnBy, valueResolved, wasResovled, linkToProp)
    ' Get the product metadata from CONFIGURATION properties
    longStatus = swCustomPropMgr1.Get6("Description", True, description, valueResolved, wasResovled, linkToProp)
    longStatus = swCustomPropMgr1.Get6("Color", True, color, valueResolved, wasResovled, linkToProp)
    
    ' Set the full drawing template path
    templatePath = "S:\Engineering\SolidWorks\Drawing Document Properties\Templates\"
    If subCategory = "DADDY" Then
        subCategory = "Daddy"
        templateFile = "scrubDaddyPartTemplate.DRWDOT"
    ElseIf subCategory = "DADDY ESSENTIAL" Then
        subCategory = "Daddy Essential"
        templateFile = "scrubDaddyEssentialPartTemplate.DRWDOT"
    ElseIf subCategory = "MOMMY" Then
        subCategory = "Mommy"
        templateFile = "scrubMommyPartTemplate.DRWDOT"
    ElseIf subCategory = "MOMMY ESSENTIAL" Then
        subCategory = "Mommy Essential"
        templateFile = "scrubMommyEssentialPartTemplate.DRWDOT"
    ElseIf subCategory = "CUSTOM CONFIGURATION" Then
        subCategory = "Dish Daddy"
        templateFile = "dishDaddyPartTemplate.DRWDOT"
    End If
            
    sheetWidth = 0.42 ' A3 paper width 420mm in landscape orientation
    sheetHeight = 0.297 ' A3 paper height 297mm in landscape orientation
    
    ' Create a new drawing
    Set swDrawing = swApp.NewDocument(templatePath & templateFile, swDwgPaperA3size, sheetWidth, sheetHeight)
    Set swDrawing = swApp.ActiveDoc
    Set swModel = swDrawing
    
    ' Set the current drawing layer to "None"
    boolStatus = swDrawing.SetCurrentLayer("")

    ' Zoom drawing sheet to maximum size in window
    swDrawing.Extension.ViewZoomToSheet

    ' Create Front View initially at left edge of page so we can measure it and
    ' determine precisely where to place it on the page
    Set swFrontView = swDrawing.CreateDrawViewFromModelView3(partName, "*Front", 0, sheetHeight / 2, 0)
    
    ' Determine material thickness based on sub-category
    If subCategory = "Daddy" Or subCategory = "Mommy" Then
        thickness = 0.0397
    ElseIf subCategory = "Daddy Essential" Or subCategory = "Mommy Essential" Then
        thickness = 0.0254
    ElseIf subCategory = "Dish Daddy" Then
        thickness = 0.02323
    End If
    
    ' Get the view bounding box position for the Front drawing view (Drawing View1)
    outlineFront = swFrontView.GetOutline
    width = outlineFront(2) - outlineFront(0)
    height = outlineFront(3) - outlineFront(1)
    viewPosFront = swFrontView.Position
    
    ' The document template apparently auto-scales depending on the size of the imported model
    ' To compensate for this in positioning schemes used throughout this macro, set a scale variable
    Set currentSheet = swDrawing.GetCurrentSheet
    sheetProperties = currentSheet.GetProperties2
    sheetScale = sheetProperties(3)

    ' Change x-position of the Front View based on the view width
    viewPosFront(0) = viewPosFront(0) + ((width / 2) + 0.03)
    swFrontView.Position = viewPosFront
    ' Update the outline coordinates since the view position changed
    outlineFront = swFrontView.GetOutline
    
    ' Create Right View by making a projected view from the Front View
    boolStatus = swDrawing.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0.1, 0.17, 0, False, 0, Nothing, 0)
    Set swRightView = swDrawing.CreateUnfoldedViewAt3(viewPosFront(0) + (width / 2) + (0.08 / sheetScale) + (thickness / 2), sheetHeight / 2, 0, False)
    viewPosRight = swRightView.Position
    outlineRight = swRightView.GetOutline
    
    ' Create Iso View
    Set swIsoView = swDrawing.CreateDrawViewFromModelView3(partName, "*Isometric", 0.37, 0.2, 0)
    viewPosIso = swIsoView.Position
    outlineIso = swIsoView.GetOutline
    isoWidth = outlineIso(2) - outlineIso(0)
    isoHeight = outlineIso(3) - outlineIso(1)
    
    ' Ensure that the Iso View is positioned to the right-most edge of the page just below the revision table
    viewPosIso(0) = sheetWidth - (0.005 * sheetScale) - (isoWidth / 2)
    viewPosIso(1) = sheetHeight - (0.015 * sheetScale) - (isoHeight / 2)
    swIsoView.Position = viewPosIso
    outlineIso = swIsoView.GetOutline
    
    ' Create isometric view label
    swIsoView.ScaleDecimal = 0.5
    isoNoteText = "ISO VIEW" & Chr(13) & "SCALE " & swIsoView.ScaleRatio(0) & ":" & swIsoView.ScaleRatio(1)
    Set isoNote = swDrawing.CreateText2(isoNoteText, viewPosIso(0), viewPosIso(1) - (isoHeight / 2) + (0.04 / sheetScale) - 0.01, 0, 0.00291, 0)
    isoNote.SetTextJustification swTextJustification_e.swTextJustificationCenter
    Set isoAnnotation = isoNote.GetAnnotation
    longStatus = isoAnnotation.SetLeader3(swNO_LEADER, swLS_SMART, True, True, True, False)
    isoAnnotation.Select3 False, selectData
    boolStatus = swDrawing.ActivateView("Drawing View3")
    boolStatus = swDrawing.Extension.SelectByID2("", "", viewPosIso(0), viewPosIso(1), 0, True, 0, Nothing, 0)
    swDrawing.AttachAnnotation swAttachAnnotationOption_e.swAttachAnnotationOption_View
    swModel.ClearSelection2 True
        
    ' Make the bounding boxes visible throughout the drawing
    boolStatus = swDrawing.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewDispGlobalBBox, True)

    ' Hide the bounding boxes for projected side view (Drawing View2) and isometric view (Drawing View3)
    boolStatus = swDrawing.Extension.SelectByID2("Bounding Box@" & partName & "-3@Drawing View3", "BBOXSKETCH", 0, 0, 0, False, 0, Nothing, 0)
    swDrawing.BlankSketch
        
    ' Add row to the revision table and pre-populate values into the relevant cells.
    Set currentSheet = swDrawing.GetCurrentSheet()
    Set myRevisionTable = currentSheet.RevisionTable
    longStatus = myRevisionTable.AddRevision("")
    myRevisionTable.Text2(2, 2, True) = "Initial Release"
    myRevisionTable.Text2(2, 3, True) = drawnBy & " / " & Format(Date, "ddmmmyyyy")
    myRevisionTable.Text2(2, 4, True) = "J. Sobel / " & Format(Date, "ddmmmyyyy")
    
    ' Add hyperlink into the drawing to link to the artwork file
    artworkFileFullPath = newProductSetup.artworkFilePath.Text
    If artworkFileFullPath <> "" Then
        'Debug.Print "artworkFullPath: " & artworkFileFullPath
        artworkFileRelativePath = GetRelativePath(saveDirectory, artworkFileFullPath)
        'Debug.Print "saveDirectory:   " & saveDirectory
        'Debug.Print "relativePath: " & artworkFileRelativePath
        artworkFileName = Right(artworkFileFullPath, Len(artworkFileFullPath) - InStrRev(artworkFileFullPath, "\"))
        Set artworkFileNote = swDrawing.CreateText2("(" & artworkFileName & ")", 0.036, 0.024, 0, 0.00291, 0)
        'Set artworkFileNote = swDrawing.CreateText2("(file://" & artworkFileRelativePath & ")", 0.036, 0.024, 0, 0.00291, 0)
        artworkFileNote.SetTextJustification swTextJustification_e.swTextJustificationLeft
        Set artworkFileAnnotation = artworkFileNote.GetAnnotation
        longStatus = artworkFileAnnotation.SetLeader3(swNO_LEADER, swLS_SMART, True, True, True, False)
        'boolStatus = artworkFileNote.SetHyperlinkText("FILE://" & artworkFileRelativePath)
        'boolStatus = artworkFileNote.SetHyperlinkText(artworkFileRelativePath)
        boolStatus = artworkFileNote.SetHyperlinkText(artworkFileFullPath) ' Full path
    End If
    If boolStatus = True Then
        'Debug.Print "Hyperlink: " & artworkFileNote.GetHyperlinkText
    End If
    
    ' Add Flextexture material/color callout flag to Drawing View2
    Set matlNote1 = swDrawing.CreateText2("<FlagNotes#NPer-Flag-1>", outlineRight(0) - 0.1, viewPosRight(1), 0, 0.00291, 0)
    Set matlNoteAnnotation = matlNote1.GetAnnotation
    longStatus = matlNoteAnnotation.SetLeader3(swLeaderStyle_e.swBENT, swLS_RIGHT, True, False, False, False)
    boolStatus = matlNoteAnnotation.SetPosition(outlineRight(0) - 0.02, viewPosRight(1) + 0.015, 0)
    boolStatus = matlNoteAnnotation.SetTextFormat(0, True, matlNoteFormat)
    boolStatus = matlNoteAnnotation.SetLeaderAttachmentPointAtIndex(0, viewPosRight(0) - (thickness / (2 * sheetScale)), viewPosRight(1), 0)

    ' Add material/color callout flag to Drawing View2 if this is a Dish Daddy, Scrub Mommy, or Scrub Mommy Essential product
    If subCategory = "Mommy" Or subCategory = "Mommy Essential" Or subCategory = "Dish Daddy" Then
        Set matlNote2 = swDrawing.CreateText2("<FlagNotes#NPer-Flag-2>", outlineRight(0) - 0.1, viewPosRight(1), 0, 0.00291, 0)
        Set matlNoteAnnotation = matlNote2.GetAnnotation
        longStatus = matlNoteAnnotation.SetLeader3(swLeaderStyle_e.swBENT, swLS_LEFT, True, False, False, False)
        boolStatus = matlNoteAnnotation.SetPosition(outlineRight(2) + 0.01, viewPosRight(1) + 0.015, 0)
        boolStatus = matlNoteAnnotation.SetTextFormat(0, True, matlNoteFormat)
        boolStatus = matlNoteAnnotation.SetLeaderAttachmentPointAtIndex(0, viewPosRight(0) + ((thickness / 2) / sheetScale), viewPosRight(1), 0)
    End If
    
    ' Add Velcro loop material/color callout flag to Drawing View2 if this is a Dish Daddy product
    If subCategory = "Dish Daddy" Then
        Set matlNote2 = swDrawing.CreateText2("<FlagNotes#NPer-Flag-3>", outlineRight(0) - 0.1, viewPosRight(1), 0, 0.00291, 0)
        Set matlNoteAnnotation = matlNote2.GetAnnotation
        longStatus = matlNoteAnnotation.SetLeader3(swLeaderStyle_e.swBENT, swLS_LEFT, True, False, False, False)
        boolStatus = matlNoteAnnotation.SetPosition(outlineRight(2) + 0.01, outlineRight(3) + 0.01, 0)
        boolStatus = matlNoteAnnotation.SetTextFormat(0, True, matlNoteFormat)
        boolStatus = matlNoteAnnotation.SetLeaderAttachmentPointAtIndex(0, viewPosRight(0), outlineRight(3) - 0.01, 0)
        selectionEnum = matlNoteAnnotation.SetArrowHeadStyleAtIndex(0, swArrowStyle_e.swDOT_ARROWHEAD)
    End If
        
    ' Disable "input dimension value" option
    swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swInputDimValOnCreate, False
        
    ' Create overall thickness dimension
    ' First try selecting vertices and if that doesn't work, try selecting external sketch points
    boolStatus = swDrawing.ActivateSheet("Sheet1")
    boolStatus = swDrawing.ActivateView("Drawing View2")
    boolStatus = swDrawing.Extension.SelectByRay(outlineRight(0), outlineRight(3), 1, 0, 0, -1, 0.01, swSelVERTICES, False, 0, swSelectOptionDefault)  ' Select vertex at top left corner of view
    If boolStatus = False Then
        boolStatus = swDrawing.Extension.SelectByID2("", "swSelEXTSKETCHPOINTS", outlineRight(0) + 0.005, outlineRight(3) - 0.005, -(thickness / 2), False, 0, Nothing, swSelectOptionDefault)
    End If
    boolStatus = swDrawing.Extension.SelectByRay(outlineRight(2), outlineRight(3), 1, 0, 0, -1, 0.01, swSelVERTICES, True, 0, swSelectOptionDefault) ' Select vertex at top right corner of view
    If boolStatus = False Then
        boolStatus = swDrawing.Extension.SelectByID2("", "swSelEXTSKETCHPOINTS", outlineRight(2) - 0.005, outlineRight(3) - 0.005, -(thickness / 2), True, 0, Nothing, swSelectOptionDefault)
    End If
    Set myDisplayDim = swDrawing.AddHorizontalDimension2(outlineRight(0) - 0.01, outlineRight(3) + 0.01, 0) ' place dimension at top, center of view
    ' Error handling in case the entities could not be selected for whatever reason
    If myDisplayDim Is Nothing Then
        response = MsgBox("Couldn't create this dimension:" & vbCrLf & vbCrLf & "Overall Thickness", vbOKOnly Or vbExclamation, "Oops! You'll need to check this later")
    Else
        myDisplayDim.SetPrecision3 1, 1, 0, 0
    End If
    boolStatus = swDrawing.Extension.SelectByID2("RD1@Drawing View2", "DIMENSION", 0, 0, 0, False, 0, Nothing, swSelectOptionDefault)
    If subCategory = "Daddy" Or subCategory = "Daddy Essential" Then
        boolStatus = swModel.EditDimensionProperties2(swTolSYMMETRIC, 0.001, 0, "", "", False, 1, 2, True, 12, 12, "", "", True, "", "", False) ' add symmeetric tolerance
    ElseIf subCategory = "Mommy" Or subCategory = "Mommy Essential" Or subCategory = "Dish Daddy" Then
        boolStatus = swModel.EditDimensionProperties2(swTolSYMMETRIC, 0.001, 0, "", "", False, 1, 2, True, 12, 12, "(", ")", True, "", "", False) ' add symmetric tolerance and make reference
    End If
    swModel.ClearSelection2 True
    
    ' Create thickness dimensions for Scrub Mommy and Scrub Mommy Essential products
    If subCategory = "Mommy" Or subCategory = "Mommy Essential" Then
        ' Flextexture thickness dimension
        boolStatus = swDrawing.ActivateSheet("Sheet1")
        boolStatus = swDrawing.ActivateView("Drawing View2")
        boolStatus = swDrawing.Extension.SelectByRay(outlineRight(0), outlineRight(3), 1, 0, 0, -1, 0.01, swSelVERTICES, False, 0, swSelectOptionDefault) ' Select vertex at top left corner of view
        If boolStatus = False Then
            boolStatus = swDrawing.Extension.SelectByID2("", "swSelEXTSKETCHPOINTS", outlineRight(0) + 0.005, outlineRight(3) - 0.005, -(thickness / 2), False, 0, Nothing, swSelectOptionDefault)
        End If
        boolStatus = swDrawing.Extension.SelectByRay(viewPosRight(0), outlineRight(3), 1, 0, 0, -1, 0.01, swSelVERTICES, True, 0, swSelectOptionDefault) ' Select vertex in center of view
        If boolStatus = False Then
            boolStatus = swDrawing.Extension.SelectByRay(viewPosRight(0), viewPosRight(1), 1, 0, 0, -1, 0.005, swSelEDGES, True, 0, swSelectOptionDefault) ' Select vertex in center of view
        End If
        Set myDisplayDim = swDrawing.AddHorizontalDimension2(outlineRight(0) - 0.01, outlineRight(1), 0)
        ' Error handling in case the entities could not be selected for whatever reason
        If myDisplayDim Is Nothing Then
            response = MsgBox("Couldn't create this dimension:" & vbCrLf & vbCrLf & "Flextexture Thickness", vbOKOnly Or vbExclamation, "Oops! You'll need to check this later")
        Else
            myDisplayDim.SetPrecision3 2, 2, 2, 2
        End If
        boolStatus = swDrawing.Extension.SelectByID2("RD2@Drawing View2", "DIMENSION", 0, 0, 0, False, 0, Nothing, swSelectOptionDefault)
        boolStatus = swModel.EditDimensionProperties2(swTolSYMMETRIC, 0.0005, 0, "", "", False, 2, 2, True, 12, 12, "", "", True, "", "", False)
        swModel.ClearSelection2 True
        'Resofoam thickness dimension
        boolStatus = swDrawing.Extension.SelectByRay(outlineRight(2), outlineRight(3), 1, 0, 0, -1, 0.01, swSelVERTICES, False, 0, swSelectOptionDefault) ' Select vertex at top right corner of view
        If boolStatus = False Then
            boolStatus = swDrawing.Extension.SelectByID2("", "swSelEXTSKETCHPOINTS", outlineRight(2) - 0.005, outlineRight(3) - 0.005, -(thickness / 2), True, 0, Nothing, swSelectOptionDefault)
        End If
        boolStatus = swDrawing.Extension.SelectByRay(viewPosRight(0), outlineRight(3), 1, 0, 0, -1, 0.01, swSelVERTICES, False, 0, swSelectOptionDefault) ' Select vertex in center of view
        If boolStatus = False Then
            boolStatus = swDrawing.Extension.SelectByRay(viewPosRight(0), viewPosRight(1), 1, 0, 0, -1, 0.005, swSelEDGES, True, 0, swSelectOptionDefault) ' Select vertex in center of view
        End If
        Set myDisplayDim = swDrawing.AddHorizontalDimension2(outlineRight(2) + 0.01, outlineRight(1), 0)
        ' Error handling in case the entities could not be selected for whatever reason
        If myDisplayDim Is Nothing Then
            response = MsgBox("Couldn't create this dimension:" & vbCrLf & vbCrLf & "Resofoam Thickness", vbOKOnly Or vbExclamation, "Oops! You'll need to check this later")
        Else
            myDisplayDim.SetPrecision3 2, 2, 2, 2
        End If
        boolStatus = swDrawing.Extension.SelectByID2("RD3@Drawing View2", "DIMENSION", 0, 0, 0, False, 0, Nothing, swSelectOptionDefault)
        boolStatus = swModel.EditDimensionProperties2(swTolSYMMETRIC, 0.0005, 0, "", "", False, 2, 2, True, 12, 12, "", "", True, "", "", False)
        swModel.ClearSelection2 True
    End If

    ' Create thickness dimensions for Dish Daddy products
    If subCategory = "Dish Daddy" Then
        ' Flextexture thickness dimension
        boolStatus = swDrawing.ActivateSheet("Sheet1")
        boolStatus = swDrawing.ActivateView("Drawing View2")
        boolStatus = swDrawing.Extension.SelectByRay(outlineRight(0), outlineRight(3), 1, 0, 0, -1, 0.01, swSelVERTICES, False, 0, swSelectOptionDefault) ' Select vertex at top left corner of view
        If boolStatus = False Then
            boolStatus = swDrawing.Extension.SelectByID2("", "swSelEXTSKETCHPOINTS", outlineRight(0) + 0.005, outlineRight(3) - 0.005, -(thickness / 2), False, 0, Nothing, swSelectOptionDefault)
        End If
        boolStatus = swDrawing.Extension.SelectByRay(viewPosRight(0) - 0.003235, outlineRight(3), 1, 0, 0, -1, 0.01, swSelVERTICES, True, 0, swSelectOptionDefault) ' Select vertex in center of view
        If boolStatus = False Then
            boolStatus = swDrawing.Extension.SelectByRay(viewPosRight(0) - 0.003235, viewPosRight(1), 1, 0, 0, -1, 0.005, swSelEDGES, True, 0, swSelectOptionDefault) ' Select vertex in center of view
        End If
        Set myDisplayDim = swDrawing.AddHorizontalDimension2(outlineRight(0) - 0.01, outlineRight(1), 0)
        ' Error handling in case the entities could not be selected for whatever reason
        If myDisplayDim Is Nothing Then
            response = MsgBox("Couldn't create this dimension:" & vbCrLf & vbCrLf & "Flextexture Thickness", vbOKOnly Or vbExclamation, "Oops! You'll need to check this later")
        Else
            myDisplayDim.SetPrecision3 2, 2, 2, 2
        End If
        boolStatus = swDrawing.Extension.SelectByID2("RD2@Drawing View2", "DIMENSION", 0, 0, 0, False, 0, Nothing, swSelectOptionDefault)
        boolStatus = swModel.EditDimensionProperties2(swTolSYMMETRIC, 0.0005, 0, "", "", False, 2, 2, True, 12, 12, "", "", True, "", "", False)
        swModel.ClearSelection2 True
        'Resofoam thickness dimension
        boolStatus = swDrawing.Extension.SelectByRay(viewPosRight(0) - 0.003235, outlineRight(3), 1, 0, 0, -1, 0.01, swSelVERTICES, False, 0, swSelectOptionDefault) ' Select vertex at top right corner of view
        If boolStatus = False Then
            boolStatus = swDrawing.Extension.SelectByID2("", "swSelEXTSKETCHPOINTS", viewPosRight(0) - 0.003235, outlineRight(3) - 0.005, -(thickness / 2), True, 0, Nothing, swSelectOptionDefault)
        End If
        boolStatus = swDrawing.Extension.SelectByRay(viewPosRight(0) + 0.010615, outlineRight(3), 1, 0, 0, -1, 0.01, swSelVERTICES, False, 0, swSelectOptionDefault) ' Select vertex in center of view
        If boolStatus = False Then
            boolStatus = swDrawing.Extension.SelectByRay(viewPosRight(0) + 0.010615, viewPosRight(1), 1, 0, 0, -1, 0.005, swSelEDGES, True, 0, swSelectOptionDefault) ' Select vertex in center of view
        End If
        Set myDisplayDim = swDrawing.AddHorizontalDimension2(outlineRight(0) - 0.01, outlineRight(1) - 0.01, 0)
        ' Error handling in case the entities could not be selected for whatever reason
        If myDisplayDim Is Nothing Then
            response = MsgBox("Couldn't create this dimension:" & vbCrLf & vbCrLf & "Resofoam Thickness", vbOKOnly Or vbExclamation, "Oops! You'll need to check this later")
        Else
            myDisplayDim.SetPrecision3 2, 2, 2, 2
        End If
        boolStatus = swDrawing.Extension.SelectByID2("RD3@Drawing View2", "DIMENSION", 0, 0, 0, False, 0, Nothing, swSelectOptionDefault)
        boolStatus = swModel.EditDimensionProperties2(swTolSYMMETRIC, 0.0005, 0, "", "", False, 2, 2, True, 12, 12, "", "", True, "", "", False)
        swModel.ClearSelection2 True
        'Velcro Loop thickness dimension
        boolStatus = swDrawing.Extension.SelectByRay(outlineRight(2), outlineRight(3), 1, 0, 0, -1, 0.01, swSelVERTICES, False, 0, swSelectOptionDefault) ' Select vertex at top right corner of view
        If boolStatus = False Then
            boolStatus = swDrawing.Extension.SelectByID2("", "swSelEXTSKETCHPOINTS", outlineRight(2) - 0.005, outlineRight(3) - 0.005, -(thickness / 2), True, 0, Nothing, swSelectOptionDefault)
        End If
        boolStatus = swDrawing.Extension.SelectByRay(viewPosRight(0) + 0.010615, outlineRight(3), 1, 0, 0, -1, 0.01, swSelVERTICES, False, 0, swSelectOptionDefault) ' Select vertex in center of view
        If boolStatus = False Then
            boolStatus = swDrawing.Extension.SelectByRay(viewPosRight(0) + 0.010615, viewPosRight(1), 1, 0, 0, -1, 0.005, swSelEDGES, True, 0, swSelectOptionDefault) ' Select vertex in center of view
        End If
        Set myDisplayDim = swDrawing.AddHorizontalDimension2(outlineRight(2) + 0.01, outlineRight(1), 0)
        ' Error handling in case the entities could not be selected for whatever reason
        If myDisplayDim Is Nothing Then
            response = MsgBox("Couldn't create this dimension:" & vbCrLf & vbCrLf & "Velcro Loop Thickness", vbOKOnly Or vbExclamation, "Oops! You'll need to check this later")
        Else
            myDisplayDim.SetPrecision3 3, 3, 3, 3
        End If
        boolStatus = swDrawing.Extension.SelectByID2("RD4@Drawing View2", "DIMENSION", 0, 0, 0, False, 0, Nothing, swSelectOptionDefault)
        boolStatus = swModel.EditDimensionProperties2(swTolNONE, 0, 0, "", "", False, 3, 2, True, 12, 12, "", "", True, "", "", False)
        swModel.ClearSelection2 True
    End If

    ' Add bounding box dimensions to Front View
    boolStatus = swDrawing.ActivateSheet("Sheet1")
    boolStatus = swDrawing.ActivateView("Drawing View1")
    ' Create the horizontal width dimension for the product (bounding box)
    boolStatus = swDrawing.Extension.SelectByID2("", "EXTSKETCHSEGMENT", viewPosFront(0), outlineFront(3) - 0.005, -(thickness / 2), False, 0, Nothing, swSelectOptionDefault)
    Set myDisplayDim = swDrawing.AddHorizontalDimension2(viewPosFront(0), outlineFront(3) + 0.01, 0)
    boolStatus = swDrawing.Extension.SelectByID2("RD1@Drawing View1", "DIMENSION", 0, 0, 0, False, 0, Nothing, swSelectOptionDefault)
    boolStatus = swModel.EditDimensionProperties2(swTolNONE, 0, 0, "", "", False, 2, 2, True, 12, 12, "(", ")", True, "", "", False) ' make reference
    ' Create the vertical height dimension for the product (bounding box)
    boolStatus = swDrawing.Extension.SelectByID2("", "EXTSKETCHSEGMENT", outlineFront(2) - 0.005, viewPosFront(1), -(thickness / 2), False, 0, Nothing, swSelectOptionDefault)
    Set myDisplayDim = swDrawing.AddVerticalDimension2(outlineFront(2) + 0.01, viewPosFront(1), 0)
    boolStatus = swDrawing.Extension.SelectByID2("RD2@Drawing View1", "DIMENSION", 0, 0, 0, False, 0, Nothing, swSelectOptionDefault)
    boolStatus = swModel.EditDimensionProperties2(swTolNONE, 0, 0, "", "", False, 2, 2, True, 12, 12, "(", ")", True, "", "", False) ' make reference

    ' Rebuild the drawing
    swDrawing.EditRebuild
    
    ' Set custom properties for this drawing file
    Set swCustomPropMgr2 = swModel.Extension.CustomPropertyManager("")
    longStatus = swCustomPropMgr2.Add3("Category", swCustomInfoType_e.swCustomInfoText, "FOAM", swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd)
    If subCategory = "Dish Daddy" Then
        longStatus = swCustomPropMgr2.Add3("Sub Category", swCustomInfoType_e.swCustomInfoText, "CUSTOM CONFIGURATION", swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd)
    Else
        longStatus = swCustomPropMgr2.Add3("Sub Category", swCustomInfoType_e.swCustomInfoText, UCase(subCategory), swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd)
    End If
    longStatus = swCustomPropMgr2.Add3("Description", swCustomInfoType_e.swCustomInfoText, description, swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd)
    longStatus = swCustomPropMgr2.Add3("Material", swCustomInfoType_e.swCustomInfoText, "", swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd)
    longStatus = swCustomPropMgr2.Add3("Color", swCustomInfoType_e.swCustomInfoText, color, swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd)
    longStatus = swCustomPropMgr2.Add3("Finish", swCustomInfoType_e.swCustomInfoText, "N/A", swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd)
    longStatus = swCustomPropMgr2.Add3("DrawnBy", swCustomInfoType_e.swCustomInfoText, newProductSetup.engineerList.Text, swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd)
    longStatus = swCustomPropMgr2.Add3("DrawnDate", swCustomInfoType_e.swCustomInfoText, Format(Date, "ddmmmyyyy"), swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd)
        
    ' Save the drawing using the CAD model name + color appended as the filename
    longStatus = swModel.Extension.SaveAs(saveDirectory & partName & "_" & Replace(color, " ", "") & ".SLDDRW", swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, longErrors, longWarnings)
    
    ' Close the drawing
    swApp.CloseDoc (swDrawing.GetPathName)
    
End Sub



' ******************************************************************************
' Get relative file path
' Harshil Patel
' July 23, 2025
' Scrub Daddy, Inc.
'
' This function compares two strings and finds where they differ by iterating
' through the characters of both strings and identifying the position of the
' first mismatch between the two.
'
' This function is used to ultimately create a relative path to the artwork
' file that can be used for the hyperlink on the drawings in the save directory.
'
' ******************************************************************************

Function GetRelativePath(saveDirectory As String, artworkFile As String) As String
    ' saveDirectory: path to directory (base directory)
    ' artworkFile: path to file (target file)
    ' Returns: relative path from saveDirectory to artworkFile
    
    Dim basePath As String
    Dim targetPath As String
    Dim baseArray() As String
    Dim targetArray() As String
    Dim I As Integer
    Dim commonIndex As Integer
    Dim artworkFile_relativePath As String
    
    ' Normalize paths by removing trailing backslashes and converting to lowercase
    basePath = LCase(Trim(saveDirectory))
    targetPath = LCase(Trim(artworkFile))
    
    If Right(basePath, 1) = "\" Then
        basePath = Left(basePath, Len(basePath) - 1)
    End If
    
    ' Split paths into arrays
    baseArray = Split(basePath, "\")
    targetArray = Split(targetPath, "\")
    
    ' Find the common path
    commonIndex = -1
    For I = 0 To UBound(baseArray)
        If I > UBound(targetArray) Then Exit For
        If baseArray(I) = targetArray(I) Then
            commonIndex = I
        Else
            Exit For
        End If
    Next I
    
    ' Build relative path
    artworkFile_relativePath = ""
    
    ' Add ".." for each directory level we need to go up from base
    For I = commonIndex + 1 To UBound(baseArray)
        If artworkFile_relativePath = "" Then
            artworkFile_relativePath = ".."
        Else
            artworkFile_relativePath = artworkFile_relativePath & "\.."
        End If
    Next I
    
    ' Add the remaining path from target
    For I = commonIndex + 1 To UBound(targetArray)
        If artworkFile_relativePath = "" Then
            artworkFile_relativePath = targetArray(I)
        Else
            artworkFile_relativePath = artworkFile_relativePath & "\" & targetArray(I)
        End If
    Next I
    
    ' Handle case where target is in the same directory or subdirectory
    If artworkFile_relativePath = "" Then
        ' Target is the same as base directory
        artworkFile_relativePath = "."
    ElseIf commonIndex = UBound(baseArray) And UBound(targetArray) > UBound(baseArray) Then
        ' Target is in a subdirectory of base
        artworkFile_relativePath = ""
        For I = UBound(baseArray) + 1 To UBound(targetArray)
            If artworkFile_relativePath = "" Then
                artworkFile_relativePath = targetArray(I)
            Else
                artworkFile_relativePath = artworkFile_relativePath & "\" & targetArray(I)
            End If
        Next I
    End If
    
    GetRelativePath = artworkFile_relativePath
End Function

' Example usage:
' Sub TestRelativePath()
'     Dim result As String
'     result = GetRelativePath("C:\Users\Documents", "C:\Users\Documents\Projects\file.txt")
'     Debug.Print result  ' Output: Projects\file.txt
'
'     result = GetRelativePath("C:\Users\Documents\Projects", "C:\Users\Documents\file.txt")
'     Debug.Print result  ' Output: ..\file.txt
'
'     result = GetRelativePath("C:\Projects\Alpha", "C:\Projects\Beta\file.txt")
'     Debug.Print result  ' Output: ..\Beta\file.txt
' End Sub

