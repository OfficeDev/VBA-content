---
title: ChartObjects.Placement Property (Excel)
keywords: vbaxl10.chm497086
f1_keywords:
- vbaxl10.chm497086
ms.prod: excel
api_name:
- Excel.ChartObjects.Placement
ms.assetid: 954e98e5-8b88-6918-3cbd-f8e982c0a47e
ms.date: 06/08/2017
---


# ChartObjects.Placement Property (Excel)

Returns or sets a  **Variant** value, containing an **[XlPlacement](xlplacement-enumeration-excel.md)** constant, that represents the way the objects are attached to the cells below them.


## Syntax

 _expression_ . **Placement**

 _expression_ A variable that represents a **ChartObjects** object.


## Example

This example sets the objects on Sheet1 to be free-floating (they neither moves nor are they sized with underlying cells).


```vb
Worksheets("Sheet1").ChartObjects.Placement = xlFreeFloating
```


