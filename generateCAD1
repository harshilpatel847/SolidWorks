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
'   c) Scrub Daddy Essentials (25.40mm thick Flextexture)
'   d) Scrub Mommy Essentials (12.70mm thick Flextexture, 12.70mm thick Resofoam)
'   e) Dish Daddy (8.38mm thick Flextexture, 13.84mm thick Resofoam, 1mm thick Velcro Loop)
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
Dim swConfig As SldWorks.Configuration
Dim fileName As String
Dim importData As SldWorks.ImportDxfDwgData
Dim subCategory As String
Dim subCategoryFormatted As String
Dim retVal As Boolean
Dim boolStatus As Boolean
Dim flextextureExtrude As Object
Dim resofoamExtrude As Object
Dim velcroLoopExtrude As Object
Dim materialDirectory As String
Dim resofoam As String
Dim flextexture As String
Dim velcroloop As String
Dim swRenderMaterial_flex As SldWorks.RenderMaterial
Dim swRenderMaterial_reso As SldWorks.RenderMaterial
Dim swRenderMaterial_velcro As SldWorks.RenderMaterial
Dim swEntity As SldWorks.Entity
Dim swSelMgr As SldWorks.SelectionMgr
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
Dim materialID1 As Long
Dim materialID2 As Long

Sub main()

    ' User input box to get the path of the DXF file
    Set swApp = Application.SldWorks
    fileName = newProductSetup.artworkFilePath.Text
    Set importData = swApp.GetImportFileData(fileName)

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

    ' Import method
    importData.ImportMethod("") = swConst.swImportDxfDwg_ImportMethod_e.swImportDxfDwg_ImportToPartSketch

   ' Load the specified DXF/DWG file
    Set swModel = swApp.LoadFile4(fileName, "", importData, longErrors)
    importData.AddSketchConstraints("") = True
    retVal = importData.SetMergePoints("", True, 0.000002)
    importData.ImportDimensions("") = True
    importData.ImportHatch("") = False

    ' Setup materials for appearance application
    materialDirectory = "S:\Engineering\SolidWorks\Materials and Pictures\Foam Colors & Textures"
    resofoam = materialDirectory & "\" & "resofoam.p2m"
    flextexture = materialDirectory & "\" & "orange.p2m"
    velcroloop = "S:\Engineering\SolidWorks\Materials and Pictures\Velcro\dishdaddyvelcro.p2m"
    Set swRenderMaterial_flex = swModel.Extension.CreateRenderMaterial(flextexture)
    swRenderMaterial_flex.FixedAspectRatio = True
    swRenderMaterial_flex.width = 0.09
    swRenderMaterial_flex.height = 0.0675
    swRenderMaterial_flex.Emission = 0.2
    Set swRenderMaterial_reso = swModel.Extension.CreateRenderMaterial(resofoam)
    swRenderMaterial_reso.FixedAspectRatio = True
    swRenderMaterial_reso.width = 0.12
    swRenderMaterial_reso.height = 0.09
    swRenderMaterial_reso.Emission = 0.2
    Set swRenderMaterial_velcro = swModel.Extension.CreateRenderMaterial(velcroloop)
    swRenderMaterial_velcro.FixedAspectRatio = True
    swRenderMaterial_velcro.width = 0.1
    swRenderMaterial_velcro.height = 0.075
    swRenderMaterial_velcro.Emission = 0.2

    ' Create boss extrude from the sketch
    Set swModel = swApp.ActiveDoc
    Set swSelMgr = swModel.SelectionManager
    Set swConfig = swModel.GetActiveConfiguration
    boolStatus = swModel.Extension.SelectByID2("Model", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
    
    If subCategory = "Daddy" Or subCategory = "Daddy Essential" Then

        ' Boss extrude to create the Flextexture product
        Set flextextureExtrude = swModel.FeatureManager.FeatureExtrusion3(False, False, False, 0, 0, thickness / 2, thickness / 2, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)
        
        ' Apply Flextexture material to the extruded body
        'Set swRenderMaterial = swModel.Extension.CreateRenderMaterial(flextexture)
        boolStatus = swModel.Extension.SelectByID2("Boss-Extrude1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Set swEntity = swSelMgr.GetSelectedObject6(1, -1)
        boolStatus = swRenderMaterial_flex.AddEntity(swEntity)
        boolStatus = swModel.Extension.AddDisplayStateSpecificRenderMaterial(swRenderMaterial_flex, swAllDisplayState, swConfig.GetDisplayStates, materialID1, materialID2)
        
    ElseIf subCategory = "Mommy" Or subCategory = "Mommy Essential" Then

        ' Boss extrude to create the Flextexture portion of the product
        Set flextextureExtrude = swModel.FeatureManager.FeatureExtrusion3(True, False, False, 0, 0, thickness / 2, 0, False, False, False, False, 0, 0, False, False, False, False, False, True, True, 0, 0, False)
        boolStatus = swModel.Extension.SelectByID2("Model", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
        ' Boss extrude to create the Resofoam portion of the product
        Set resofoamExtrude = swModel.FeatureManager.FeatureExtrusion3(True, False, True, 0, 0, thickness / 2, 0, False, False, False, False, 0, 0, False, False, False, False, False, True, True, 0, 0, False)
    
        ' Apply flextexture and resofoam materials to the extruded bodies
        ' Set swRenderMaterial_flex = swModel.Extension.CreateRenderMaterial(flextexture)
        boolStatus = swModel.Extension.SelectByID2("Boss-Extrude1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Set swEntity = swSelMgr.GetSelectedObject6(1, -1)
        boolStatus = swRenderMaterial_flex.AddEntity(swEntity)
        boolStatus = swModel.Extension.AddDisplayStateSpecificRenderMaterial(swRenderMaterial_flex, swAllDisplayState, swConfig.GetDisplayStates, materialID1, materialID2)
        ' Set swRenderMaterial = swModel.Extension.CreateRenderMaterial(resofoam)
        boolStatus = swModel.Extension.SelectByID2("Boss-Extrude2", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Set swEntity = swSelMgr.GetSelectedObject6(1, -1)
        boolStatus = swRenderMaterial_reso.AddEntity(swEntity)
        boolStatus = swModel.Extension.AddDisplayStateSpecificRenderMaterial(swRenderMaterial_reso, swAllDisplayState, swConfig.GetDisplayStates, materialID1, materialID2)
    
    ElseIf subCategory = "Dish Daddy" Then

        ' Boss extrude to create the Flextexture portion of the product
        Set flextextureExtrude = swModel.FeatureManager.FeatureExtrusion3(True, False, False, 0, 0, 0.00838, 0, False, False, False, False, 0, 0, False, False, False, False, False, True, True, 0, 0, False)
        boolStatus = swModel.Extension.SelectByID2("Model", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
        ' Boss extrude to create the Resofoam portion of the product
        Set resofoamExtrude = swModel.FeatureManager.FeatureExtrusion3(True, False, True, 0, 0, 0.01384, 0, False, False, False, False, 0, 0, False, False, False, False, False, True, True, 0, 0, False)
        boolStatus = swModel.Extension.SelectByID2("Model", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
        ' Boss extrude to create the Resofoam portion of the product
        Set velcroLoopExtrude = swModel.FeatureManager.FeatureExtrusion3(True, False, True, 0, 0, 0.001, 0, False, False, False, False, 0, 0, False, False, False, False, False, True, True, swStartOffset, 0.01384, True)
        
        ' Apply flextexture, resofoam, and velcro materials to the extruded bodies
        ' Set swRenderMaterial = swModel.Extension.CreateRenderMaterial(flextexture)
        boolStatus = swModel.Extension.SelectByID2("Boss-Extrude1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Set swEntity = swSelMgr.GetSelectedObject6(1, -1)
        boolStatus = swRenderMaterial_flex.AddEntity(swEntity)
        boolStatus = swModel.Extension.AddDisplayStateSpecificRenderMaterial(swRenderMaterial_flex, swAllDisplayState, swConfig.GetDisplayStates, materialID1, materialID2)
        ' Set swRenderMaterial = swModel.Extension.CreateRenderMaterial(resofoam)
        boolStatus = swModel.Extension.SelectByID2("Boss-Extrude2", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Set swEntity = swSelMgr.GetSelectedObject6(1, -1)
        boolStatus = swRenderMaterial_reso.AddEntity(swEntity)
        boolStatus = swModel.Extension.AddDisplayStateSpecificRenderMaterial(swRenderMaterial_reso, swAllDisplayState, swConfig.GetDisplayStates, materialID1, materialID2)
        ' Set swRenderMaterial = swModel.Extension.CreateRenderMaterial(velcroloop)
        boolStatus = swModel.Extension.SelectByID2("Boss-Extrude3", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Set swEntity = swSelMgr.GetSelectedObject6(1, -1)
        boolStatus = swRenderMaterial_velcro.AddEntity(swEntity)
        boolStatus = swModel.Extension.AddDisplayStateSpecificRenderMaterial(swRenderMaterial_velcro, swAllDisplayState, swConfig.GetDisplayStates, materialID1, materialID2)
    
    End If
    
    swModel.ClearSelection2 True
        
End Sub

