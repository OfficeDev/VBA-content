---
title: PickerDialog.Resolve Method (Office)
keywords: vbaof11.chm340006
f1_keywords:
- vbaof11.chm340006
ms.prod: office
api_name:
- Office.PickerDialog.Resolve
ms.assetid: 50b1792a-ecf0-ab66-6a9d-7f72c788d859
ms.date: 06/08/2017
---


# PickerDialog.Resolve Method (Office)

Resolves the token using the Picker Dialog and retrieves the results.


## Syntax

 _expression_. **Resolve**( **_TokenText_**, **_duplicateDlgMode_** )

 _expression_ An expression that returns a **PickerDialog** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TokenText_|Required|**String**|The text string to resolve.|
| _duplicateDlgMode_|Required|**Integer**||

### Return Value

PickerResults


## Example

Resolves entities by using the Picker Dialog object.


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
 
' Resolve the token by using Picker Dialog and get the results. 
Set objPickerResults = objPickerDialog.Resolve("johndoe", False) 

```


## See also


#### Concepts


[PickerDialog Object](pickerdialog-object-office.md)
#### Other resources


[PickerDialog Object Members](pickerdialog-members-office.md)

