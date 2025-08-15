' ******************************************************************************
' Create CAD model for specialty shape foam product from DXF/DWG
' Harshil Patel
' July 7, 2025
' Scrub Daddy, Inc.
'
' This macro generates a CAD model of the selected foam product from the provided
' artwork (DXF or DWG file)
'   a) Scrub Daddy (39.70mm thick Flextexture)
'   b) Scrub Mommy (19.85mm thick Flextexture, 19.85mm thick Resofoam)
'   c) TODO: Scrub Daddy Essentials (25.40mm thick Flextexture)
'   d) TODO: Scrub Mommy Essentials (12.70mm thick Flextexture, 12.70mm thick Resofoam)
'
' The generated model includes datums and features that will be used for
' dimensioning in the engineering drawing
'
' The model is saved based on a pre-defined naming scheme and data from the UserForm
'
' ******************************************************************************

Option Explicit

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swModelDocExt As SldWorks.ModelDocExtension
Dim fileName As String
Dim importData As SldWorks.ImportDxfDwgData
Dim subCategory As String
Dim subCategoryFormatted As String
Dim retVal As Boolean
Dim boolStatus As Boolean
Dim flextextureExtrude As Object
Dim resofoamExtrude As Object
Dim swRefPlane As SldWorks.RefPlane
Dim swFeatMgr As SldWorks.FeatureManager
Dim swBoundingBoxFeat As SldWorks.Feature
Dim swBoundingBoxFeatDef As SldWorks.BoundingBoxFeatureData
Dim featData As SldWorks.BoundingBoxFeatureData
Dim selectedPath As String
Dim newProductName As String
Dim longStatus As Long
Dim longErrors As Long, longWarnings As Long
Dim width As Double
Dim height As Double
Dim thickness As Double
Dim boundingBoxArray As Variant

Sub main()

    ' User input box to get the path of the DXF file
    Set swApp = Application.SldWorks
    fileName = newProductSetup.artworkFilePath.Text
    Set swModel = swApp.ActiveDoc

    ' Identify the product sub-category (Scrub Daddy or Scrub Mommy, regular or Essentials)
    If newProductSetup.scrubDaddyOption.Value = True Then
        subCategory = "Daddy"
        thickness = 0.0397
    ElseIf newProductSetup.scrubMommyOption.Value = True Then
        subCategory = "Mommy"
        thickness = 0.0397
    ElseIf newProductSetup.scrubDaddyEssentialOption.Value = True Then
        subCategory = "Daddy Essential"
        thickness = 0.0254
    ElseIf newProductSetup.scrubMommyEssentialOption.Value = True Then
        subCategory = "Mommy Essential"
        thickness = 0.0254
    ElseIf newProductSetup.dishDaddyOption.Value = True Then
        subCategory = "Dish Daddy"
        thickness = 0.02323
    End If
        
    ' Create bounding box to be used in drawing to establish overall height, width
    Set swFeatMgr = swModel.FeatureManager
    Set swBoundingBoxFeatDef = swFeatMgr.CreateDefinition(swConst.swFmBoundingBox)
    boolStatus = swModel.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
    swBoundingBoxFeatDef.ReferenceFaceOrPlane = swConst.swGlobalBoundingBoxFitOptions_e.swBoundingBoxType_CustomPlane
    Set swBoundingBoxFeat = swFeatMgr.CreateFeature(swBoundingBoxFeatDef)
    swModel.ClearSelection2 True

    ' Save the part using a name provided by the user.
    ' The new filename will take the form of scrubDaddy_XXXXX, scrubDaddyEssentials_XXXXX, etc. where XXXXX is the name provided
    selectedPath = newProductSetup.saveDirectory.Text
    newProductName = newProductSetup.productName.Text
    subCategoryFormatted = Replace(subCategory, " ", "")
    If subCategory = "Dish Daddy" Then
        subCategoryFormatted = "dishDaddy"
        longStatus = swModel.Extension.SaveAs(selectedPath & "\" & subCategoryFormatted & "_" & newProductName & ".SLDPRT", swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, longErrors, longWarnings)
    Else
        longStatus = swModel.Extension.SaveAs(selectedPath & "\scrub" & subCategoryFormatted & "_" & newProductName & ".SLDPRT", swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, longErrors, longWarnings)
    End If
End Sub
