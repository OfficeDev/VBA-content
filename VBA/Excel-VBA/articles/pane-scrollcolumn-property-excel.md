---
title: Pane.ScrollColumn Property (Excel)
keywords: vbaxl10.chm360076
f1_keywords:
- vbaxl10.chm360076
ms.prod: excel
api_name:
- Excel.Pane.ScrollColumn
ms.assetid: 47165fe4-299d-8863-708f-9db22ef52ed1
ms.date: 06/08/2017
---


# Pane.ScrollColumn Property (Excel)

Returns or sets the number of the leftmost column in the pane or window. Read/write  **Long** .


## Syntax

 _expression_ . **ScrollColumn**

 _expression_ A variable that represents a **Pane** object.


## Remarks

If the window is split, the  **ScrollColumn** property of the **[Window](window-object-excel.md)** object refers to the upper-left pane. If the panes are frozen, the **ScrollColumn** property of the **Window** object excludes the frozen areas.


## See also


#### Concepts


[Pane Object](pane-object-excel.md)

