---
title: Workbook.ConnectionsDisabled Property (Excel)
keywords: vbaxl10.chm199257
f1_keywords:
- vbaxl10.chm199257
ms.prod: EXCEL
api_name:
- Excel.Workbook.ConnectionsDisabled
ms.assetid: afd53cc5-12d8-4b22-3186-1359c14f662e
---


# Workbook.ConnectionsDisabled Property (Excel)

Disables the external connections or links in the workbook. Read-only


## Syntax

 _expression_ . **ConnectionsDisabled**

 _expression_ A variable that represents a **Workbook** object.


### Return Value

Boolean


## Example

Disables the external link when the workbook is opened.


```vb
Private Sub Workbook_Open() 
 ThisWorkbook.ConnectionsDisabled 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

