---
title: Worksheet.Deactivate Event (Excel)
keywords: vbaxl10.chm502077
f1_keywords:
- vbaxl10.chm502077
ms.prod: excel
api_name:
- Excel.Worksheet.Deactivate
ms.assetid: 3f66b86b-d0f0-bdc0-594c-3eb9faa44ff2
ms.date: 06/08/2017
---


# Worksheet.Deactivate Event (Excel)

Occurs when the chart, worksheet, or workbook is deactivated.


## Syntax

 _expression_ . **Deactivate**

 _expression_ A variable that represents a **Worksheet** object.


### Return Value

nothing


## Example

This example arranges all open windows when the workbook is deactivated.


```vb
Private Sub Workbook_Deactivate() 
 Application.Windows.Arrange xlArrangeStyleTiled 
End Sub
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

