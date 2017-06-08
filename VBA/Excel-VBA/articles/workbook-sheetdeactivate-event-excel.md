---
title: Workbook.SheetDeactivate Event (Excel)
keywords: vbaxl10.chm503089
f1_keywords:
- vbaxl10.chm503089
ms.prod: excel
api_name:
- Excel.Workbook.SheetDeactivate
ms.assetid: befde22b-69ce-c34f-2b9e-da5e026972e3
ms.date: 06/08/2017
---


# Workbook.SheetDeactivate Event (Excel)

Occurs when any sheet is deactivated.


## Syntax

 _expression_ . **SheetDeactivate**( **_Sh_** , )

 _expression_ An expression that returns a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The sheet. Can be a  **[Chart](chart-object-excel.md)** or **[Worksheet](worksheet-object-excel.md)** object.|

## Example

This example displays the name of each deactivated sheet.


```vb
Private Sub Workbook_SheetDeactivate(ByVal Sh As Object) 
 MsgBox Sh.Name 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

