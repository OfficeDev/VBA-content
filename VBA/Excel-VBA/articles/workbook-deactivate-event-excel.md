---
title: Workbook.Deactivate Event (Excel)
keywords: vbaxl10.chm503075
f1_keywords:
- vbaxl10.chm503075
ms.prod: excel
api_name:
- Excel.Workbook.Deactivate
ms.assetid: 6bd5411c-ac43-95cf-6755-49780ac765e9
ms.date: 06/08/2017
---


# Workbook.Deactivate Event (Excel)

Occurs when the chart, worksheet, or workbook is deactivated.


## Syntax

 _expression_ . **Deactivate**

 _expression_ A variable that represents a **Workbook** object.


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


[Workbook Object](workbook-object-excel.md)

