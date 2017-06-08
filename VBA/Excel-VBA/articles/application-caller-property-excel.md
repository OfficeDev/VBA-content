---
title: Application.Caller Property (Excel)
keywords: vbaxl10.chm133085
f1_keywords:
- vbaxl10.chm133085
ms.prod: excel
api_name:
- Excel.Application.Caller
ms.assetid: 0cfec08d-3cbc-0ab1-419a-f5b5702c3969
ms.date: 06/08/2017
---


# Application.Caller Property (Excel)

Returns information about how Visual Basic was called (for more information, see the Remarks section).


## Syntax

 _expression_ . **Caller**( **_Index_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|An index to the array. This argument is used only when the property returns an array (for more information, see the Remarks section).|

## Remarks

This property returns information about how Visual Basic was called, as shown in the following table.



|**Caller**|**Return value**|
|:-----|:-----|
|A custom function entered in a single cell|A  **Range** object specifying that cell|
|A custom function that is part of an array formula in a range of cells|A  **Range** object specifying that range of cells|
|An Auto_Open, Auto_Close, Auto_Activate, or Auto_Deactivate macro|The name of the document as text|
|A macro set by either the  **OnDoubleClick** or **OnEntry** property|The name of the chart object identifier or cell reference (if applicable) to which the macro applies|
|The  **Macro** dialog box ( **Tools** menu), or any caller not described above|The #REF! error value|

## Example

This example displays information about how Visual Basic was called.


```vb
Select Case TypeName(Application.Caller) 
 Case "Range" 
 v = Application.Caller.Address 
 Case "String" 
 v = Application.Caller 
 Case "Error" 
 v = "Error" 
 Case Else 
 v = "unknown" 
End Select 
MsgBox "caller = " &; v
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

