---
title: PickerDialog.DataHandlerId Property (Office)
keywords: vbaof11.chm340001
f1_keywords:
- vbaof11.chm340001
ms.prod: office
api_name:
- Office.PickerDialog.DataHandlerId
ms.assetid: 6c494116-74a2-1fdc-bc1c-033191adfca1
ms.date: 06/08/2017
---


# PickerDialog.DataHandlerId Property (Office)

Sets or gets the GUID of the Picker Dialog data handler component. Read/write


## Syntax

 _expression_. **DataHandlerId**

 _expression_ An expression that returns a **PickerDialog** object.


## Remarks

You must specify  **DataHandlerID** before invoking the Picker Dialog.


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

