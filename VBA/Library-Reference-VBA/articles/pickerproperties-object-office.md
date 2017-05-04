---
title: PickerProperties Object (Office)
keywords: vbaof11.chm337000
f1_keywords:
- vbaof11.chm337000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.PickerProperties
ms.assetid: 368e2b17-1b4f-484e-483f-53c7cd16a444
---


# PickerProperties Object (Office)

A collection of  **PickerProperty** objects.


## Remarks

Each  **PickerProperty** object is a Name(ID)/Value pair for passing option values to a PickerDialog object. You can get a **PickerProperties** collection object through the **Properties** property of **PickerDialog** object.


## Example

The following code sets the Picker Dialog properties and then displays the Picker Dialog.


```vb
Dim objPickerDialog As PickerDialog 
Dim objPickerProperties As PickerProperties 
Dim objPickerProperty As PickerProperty 
Dim objPickerExistingResults As PickerResults 
Dim objPickerExistingResult As PickerResult 
Dim objPickerResults As PickerResults 
 
 
' Configure the Picker Dialog properties. 
Set objPickerDialog = Application.PickerDialog 
objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}" 
objPickerDialog.Title = "Sample Picker Dialog" 
Set objPickerProperties = objPickerDialog.Properties 
Set objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", msoPickerFieldtypeText) 
Set objPickerExistingResults = objPickerDialog.CreatePickerResults 
Set objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User") 
 
' Show the Picker Dialog and get the results. 
Set objPickerResults = objPickerDialog.Show(True, objPickerExistingResult)
```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

