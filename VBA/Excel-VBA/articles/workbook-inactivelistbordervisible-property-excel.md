---
title: Workbook.InactiveListBorderVisible Property (Excel)
keywords: vbaxl10.chm199229
f1_keywords:
- vbaxl10.chm199229
ms.prod: excel
api_name:
- Excel.Workbook.InactiveListBorderVisible
ms.assetid: a6259862-9a29-f3a5-498f-633f51ec10e6
ms.date: 06/08/2017
---


# Workbook.InactiveListBorderVisible Property (Excel)

A  **Boolean** value that specifies whether list borders are visible when a list is not active. Returns **True** if the border is visible. Read/write **Boolean** .


## Syntax

 _expression_ . **InactiveListBorderVisible**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

Setting this property will affect all the lists that are on the worksheet.


## Example

The following example hides the borders of inactive lists in the workbook.


```vb
Sub HideListBorders() 
 
 ActiveWorkbook.InactiveListBorderVisible = False 
 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

