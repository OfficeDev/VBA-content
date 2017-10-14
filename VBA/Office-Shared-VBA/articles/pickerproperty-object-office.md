---
title: PickerProperty Object (Office)
keywords: vbaof11.chm336000
f1_keywords:
- vbaof11.chm336000
ms.prod: office
api_name:
- Office.PickerProperty
ms.assetid: fd3702fe-bf03-f22c-78c2-ac6c47a1d028
ms.date: 06/08/2017
---


# PickerProperty Object (Office)

Represents an object for passing a custom property. p


## Example

The following code sets the Picker Dialog properties and then displays the Picker dialog.


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


## Properties



|**Name**|
|:-----|
|[Application](pickerproperty-application-property-office.md)|
|[Creator](pickerproperty-creator-property-office.md)|
|[Id](pickerproperty-id-property-office.md)|
|[Type](pickerproperty-type-property-office.md)|
|[Value](pickerproperty-value-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
