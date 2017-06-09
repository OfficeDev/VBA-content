---
title: PickerDialog.Properties Property (Office)
keywords: vbaof11.chm340003
f1_keywords:
- vbaof11.chm340003
ms.prod: office
api_name:
- Office.PickerDialog.Properties
ms.assetid: 053b5d62-9d9a-68ed-c7ed-cf4df7053ecc
ms.date: 06/08/2017
---


# PickerDialog.Properties Property (Office)

Returns the ** PickerProperties** object to specify custom properties for data handler component. Read-only


## Syntax

 _expression_. **Properties**

 _expression_ An expression that returns a **PickerDialog** object.


## Remarks

The properties of the  **PickerProperties** object will be passed to the data handler.


## Example

The following code sets various Picker Dialog properties and retrieves the results.


```
Dim objPickerDialog As PickerDialog 
Dim objPickerProperties As PickerProperties 
 
Set objPickerDialog = Application.PickerDialog 
objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}" 
objPickerDialog.Title = "Sample Picker Dialog" 
Set objPickerProperties = objPickerDialog.Properties 
Set objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", msoPickerFieldtypeText) 
 
' Show the Picker Dialog with no existing result. 
Set objPickerResults = objPickerDialog.Show(True) 

```


## See also


#### Concepts


[PickerDialog Object](pickerdialog-object-office.md)
#### Other resources


[PickerDialog Object Members](pickerdialog-members-office.md)

