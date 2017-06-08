---
title: PickerProperties.Add Method (Office)
keywords: vbaof11.chm337003
f1_keywords:
- vbaof11.chm337003
ms.prod: office
api_name:
- Office.PickerProperties.Add
ms.assetid: a52c9607-1b0a-c37e-a3af-dc0550c64deb
ms.date: 06/08/2017
---


# PickerProperties.Add Method (Office)




## Syntax

 _expression_. **Add**( **_Id_**, **_Value_**, **_Type_** )

 _expression_ An expression that returns a **PickerProperties** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Id_|Required|**String**|Key name of the property.|
| _Value_|Required|**String**|Value of the property.|
| _Type_|Required|**MsoPickerField**|Type of the property.|

### Return Value

PickerProperty


## Example

The following code sets various properties of the  **PickerDialog** object.


```
Dim objPickerDialog As PickerDialog 
Dim objPickerProperties As PickerProperties 
 
' Configure Picker Dialog properties. 
Set objPickerDialog = Application.PickerDialog 
objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}" 
objPickerDialog.Title = "Sample Picker Dialog" 
Set objPickerProperties = objPickerDialog.Properties 
Set objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", msoPickerFieldtypeText) 

```


## See also


#### Concepts


[PickerProperties Object](pickerproperties-object-office.md)
#### Other resources


[PickerProperties Object Members](pickerproperties-members-office.md)

