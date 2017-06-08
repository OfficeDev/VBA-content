---
title: Application.SheetChange Event (Excel)
keywords: vbaxl10.chm504080
f1_keywords:
- vbaxl10.chm504080
ms.prod: excel
api_name:
- Excel.Application.SheetChange
ms.assetid: 0b06ad02-52c0-f0a3-c827-b7e51aecc81c
ms.date: 06/08/2017
---


# Application.SheetChange Event (Excel)

Occurs when cells in any worksheet are changed by the user or by an external link.


## Syntax

 _expression_ . **SheetChange**( **_Sh_** , **_Target_** )

 _expression_ An expression that returns a **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|A  **[Worksheet](worksheet-object-excel.md)** object that represents the sheet.|
| _Target_|Required| **Range**|The changed range.|

## Remarks

This event doesn't occur on chart sheets.


## Example

This example runs when any worksheet is changed.


```vb
Private Sub Workbook_SheetChange(ByVal Sh As Object, _ 
 ByVal Source As Range) 
 ' runs when a sheet is changed 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

