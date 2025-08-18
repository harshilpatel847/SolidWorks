# SolidWorks
Macros and Scripts for use with SolidWorks


## Check and Approve Macro
This macro updates the Solidworks drawing by updating the revision table and title block to include approving manager's name and the current date


## Add Mass Property Macro
This macro sets the default unit system for the currently active part to be MMGS. The number of significant figures for the mass properties is set to 6 so it's more appropriate for small parts with low densities (when using grams per millimeter cubed).


## General Export PDF Macro
This macro performs a 'Save As' operation on an open and active drawing document to export as a PDF using default settings. 


## New Product Setup Macro
Upon executing the script a UserForm is launched that gathers the necessary information for generating the appropriate CAD model and drawings. All files are named based on specifications and saved per the UserForm.\
<img height="400" alt="Screenshot 2025-07-22 124710" src="https://github.com/user-attachments/assets/d6720107-684a-4e69-957f-965064d977e8" />
<img height="400" alt="Screenshot 2025-07-22 124745" src="https://github.com/user-attachments/assets/d10d8ab4-5bf6-4997-a92a-612a94b53b49" />
<img height="400" alt="Screenshot 2025-07-22 124806" src="https://github.com/user-attachments/assets/3cfefffb-d5d6-4f9f-aa06-ff35c4c3c4c9" />\

This macro uses three procedures to generate the appropriate engineering documentation:
### 1. generateCAD Procedure
This macro generates a CAD model of the selected foam product from the provided artwork (DXF or DWG file).\
1. Scrub Daddy (39.70mm thick Flextexture)
2. Scrub Mommy (19.85mm thick Flextexture, 19.85mm thick Resofoam)
3. Scrub Daddy Essentials (25.40mm thick Flextexture)
4. Scrub Mommy Essentials (12.70mm thick Flextexture, 12.70mm thick Resofoam)\
5. Dish Daddy (Flextexture, Resofoam, and Velcro loop)\
The model is saved based on a pre-defined naming scheme and data from the UserForm
### 2. generateConfigs Procedure
This macro generates several custom configurations and adds three custom properties to each:
- Property 1: Part Number
- Property 2: Description
- Property 3: Color\
Additional properties are added to the file that Bild PDM will be able to read and apply to the file for document control.\
The number, names, and property values of the configurations are driven by selections made of page 2 of the UserForm.
### 3. generateDrawings Procedure
This macro acts on a saved part to generate a specialty foam shape drawing. The generated drawing uses a template based on the part's sub-category (Either a Scrub Mommy product or Scrub Daddy product).
