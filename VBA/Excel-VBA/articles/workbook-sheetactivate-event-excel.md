---
title: Workbook.SheetActivate Event (Excel)
keywords: vbaxl10.chm503088
f1_keywords:
- vbaxl10.chm503088
ms.prod: excel
api_name:
- Excel.Workbook.SheetActivate
ms.assetid: 2a7c05c3-5b66-8012-5ac5-981dcfc7f947
ms.date: 06/08/2017
---


# Workbook.SheetActivate Event (Excel)

Occurs when any sheet is activated.


## Syntax

 _expression_ . **SheetActivate**( **_Sh_** )

 _expression_ An expression that returns a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The activated sheet. Can be a  **[Chart](chart-object-excel.md)** or **[Worksheet](worksheet-object-excel.md)** object.|

## Example

This example displays the name of each activated sheet.


```vb
Private Sub Workbook_SheetActivate(ByVal Sh As Object) 
 MsgBox Sh.Name 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

