---
title: CustomProperties.Item Property (Excel)
keywords: vbaxl10.chm680076
f1_keywords:
- vbaxl10.chm680076
ms.prod: excel
api_name:
- Excel.CustomProperties.Item
ms.assetid: f2b9890b-2a25-e192-323b-dca72b461229
ms.date: 06/08/2017
---


# CustomProperties.Item Property (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **CustomProperties** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number of the object.|

## Example

The following example demonstrates this feature. In this example, Microsoft Excel adds identifier information to the active worksheet and returns the name and value to the user.


```vb
Sub CheckCustomProperties() 
 
 Dim wksSheet1 As Worksheet 
 
 Set wksSheet1 = Application.ActiveSheet 
 
 ' Add metadata to worksheet. 
 wksSheet1.CustomProperties.Add _ 
 Name:="Market", Value:="Nasdaq" 
 
 ' Display metadata. 
 With wksSheet1.CustomProperties.Item(1) 
 MsgBox .Name &; vbTab &; .Value 
 End With 
 
End Sub
```


## See also


#### Concepts


[CustomProperties Object](customproperties-object-excel.md)

