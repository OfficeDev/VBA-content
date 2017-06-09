---
title: PickerResults Object (Office)
keywords: vbaof11.chm339000
f1_keywords:
- vbaof11.chm339000
ms.prod: office
api_name:
- Office.PickerResults
ms.assetid: c0e2e097-021b-7ed4-2f94-8204c849bc17
ms.date: 06/08/2017
---


# PickerResults Object (Office)

A collection of  **PickerResult** objects.


## Remarks

Each  **PickerResult** object represents a resolved or selected item data.


## Example

The following code displays the Picker Dialog, gets results, and then enumerates those results.


```
Dim objPickerDialog As PickerDialog 
Dim objPickerProperties As PickerProperties 
Dim objPickerProperty As PickerProperty 
Dim objPickerExistingResults As PickerResults 
Dim objPickerExistingResults As PickerResult 
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
 
' Enumerate the results. 
For index = 1 To objPickerResults.Count-1 
 Debug.Print objPickerResults.Item(index).Id 
 Debug.Print objPickerResults.Item(index).DisplayName 
 Debug.Print objPickerResults.Item(index).Type 
 Debug.Print objPickerResults.Item(index).SIPId 
Next 

```


## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
[PickerResults Object Members](http://msdn.microsoft.com/library/6b6ec287-4d88-cc7d-7cfa-f641b1481bbe%28Office.15%29.aspx)
