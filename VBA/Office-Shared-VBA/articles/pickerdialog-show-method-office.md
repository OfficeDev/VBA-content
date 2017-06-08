---
title: PickerDialog.Show Method (Office)
keywords: vbaof11.chm340005
f1_keywords:
- vbaof11.chm340005
ms.prod: office
api_name:
- Office.PickerDialog.Show
ms.assetid: 3073defe-4585-816d-6b86-9959cce4655f
ms.date: 06/08/2017
---


# PickerDialog.Show Method (Office)

Displays the Picker Dialog with already specified data handler and given options.


## Syntax

 _expression_. **Show**( **_IsMultiSelect_**, **_ExistingResults_** )

 _expression_ An expression that returns a **PickerDialog** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _IsMultiSelect_|Optional|**Boolean**|Specifies whether the Picker Dialog user interface provides multiple item selection functions.|
| _ExistingResults_|Optional|**PickerResults**|Contains existing ** PickerResults** in Picker Dialog user interface. These results are displayed in the selected item control.|

### Return Value

PickerResults


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

