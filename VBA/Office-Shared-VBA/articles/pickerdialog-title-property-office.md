---
title: PickerDialog.Title Property (Office)
keywords: vbaof11.chm340002
f1_keywords:
- vbaof11.chm340002
ms.prod: office
api_name:
- Office.PickerDialog.Title
ms.assetid: 76531e47-91a4-4d82-7825-ab900c5bf8e2
ms.date: 06/08/2017
---


# PickerDialog.Title Property (Office)

Set or returns the title of a picker dialog displayed in the Picker Dialog. Read/write


## Syntax

 _expression_. **Title**

 _expression_ An expression that returns a **PickerDialog** object.


## Example

The following code sets the Picker Dialog properties and then displays the Picker Dialog.


```
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


[PickerDialog Object](pickerdialog-object-office.md)
#### Other resources


[PickerDialog Object Members](pickerdialog-members-office.md)

