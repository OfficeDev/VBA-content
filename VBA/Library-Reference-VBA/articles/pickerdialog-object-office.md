---
title: PickerDialog Object (Office)
keywords: vbaof11.chm340000
f1_keywords:
- vbaof11.chm340000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.PickerDialog
ms.assetid: 279b1a6a-f09d-a0e7-89c9-aac6c581439f
---


# PickerDialog Object (Office)

Provides dialog user interface functionality of for picking people or picking data.


## Remarks

Get the  **PickerDialog** object through the **PickerDialog** property in **Application** object.


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

