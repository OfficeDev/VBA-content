---
title: PickerResult.Type Property (Office)
keywords: vbaof11.chm338003
f1_keywords:
- vbaof11.chm338003
ms.prod: office
api_name:
- Office.PickerResult.Type
ms.assetid: e7e0356a-7d21-c9f4-81f3-4ac096c5ab4f
ms.date: 06/08/2017
---


# PickerResult.Type Property (Office)

Represents the type of a  **PickerResult** object. Read/write


## Syntax

 _expression_. **Type**

 _expression_ An expression that returns a **PickerResult** object.


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


[PickerResult Object](pickerresult-object-office.md)
#### Other resources


[PickerResult Object Members](pickerresult-members-office.md)

