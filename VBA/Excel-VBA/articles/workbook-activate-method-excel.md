---
title: Workbook.Activate Method (Excel)
keywords: vbaxl10.chm199074
f1_keywords:
- vbaxl10.chm199074
ms.prod: excel
api_name:
- Excel.Workbook.Activate
ms.assetid: 628e06b3-ca3f-28cb-e0fd-e696842f69f5
ms.date: 06/08/2017
---


# Workbook.Activate Method (Excel)

Activates the first window associated with the workbook.


## Syntax

 _expression_ . **Activate**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

This method won't run any Auto_Activate or Auto_Deactivate macros that might be attached to the workbook (use the  **[RunAutoMacros](workbook-runautomacros-method-excel.md)** method to run those macros).


## Example

This example activates Book4.xls. If Book4.xls has multiple windows, the example activates the first window, Book4.xls:1.


```vb
Workbooks("BOOK4.XLS").Activate
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

