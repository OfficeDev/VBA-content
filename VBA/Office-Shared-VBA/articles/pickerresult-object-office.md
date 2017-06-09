---
title: PickerResult Object (Office)
keywords: vbaof11.chm338000
f1_keywords:
- vbaof11.chm338000
ms.prod: office
api_name:
- Office.PickerResult
ms.assetid: 5229d2ad-a32e-a864-9de4-dc651199ff58
ms.date: 06/08/2017
---


# PickerResult Object (Office)

Represents a resolved or selected item of data.


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
|[Application](pickerresult-application-property-office.md)|
|[Creator](pickerresult-creator-property-office.md)|
|[DisplayName](pickerresult-displayname-property-office.md)|
|[DuplicateResults](pickerresult-duplicateresults-property-office.md)|
|[Fields](pickerresult-fields-property-office.md)|
|[Id](pickerresult-id-property-office.md)|
|[ItemData](pickerresult-itemdata-property-office.md)|
|[SIPId](pickerresult-sipid-property-office.md)|
|[SubItems](pickerresult-subitems-property-office.md)|
|[Type](pickerresult-type-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
