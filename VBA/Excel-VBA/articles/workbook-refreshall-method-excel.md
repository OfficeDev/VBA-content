---
title: Workbook.RefreshAll Method (Excel)
keywords: vbaxl10.chm199135
f1_keywords:
- vbaxl10.chm199135
ms.prod: excel
api_name:
- Excel.Workbook.RefreshAll
ms.assetid: c1a956dc-263c-5c24-3b51-fc4af22dcd33
ms.date: 06/08/2017
---


# Workbook.RefreshAll Method (Excel)

Refreshes all external data ranges and PivotTable reports in the specified workbook.


## Syntax

 _expression_ . **RefreshAll**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

Objects that have the  **[BackgroundQuery](pivotcache-backgroundquery-property-excel.md)** property set to **True** are refreshed in the background.


## Example

This example refreshes all external data ranges and PivotTable reports in the third workbook.


```vb
Workbooks(3).RefreshAll
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

