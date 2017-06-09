---
title: Shape.CellsRowIndexU Property (Visio)
keywords: vis_sdr.chm11251950
f1_keywords:
- vis_sdr.chm11251950
ms.prod: visio
api_name:
- Visio.Shape.CellsRowIndexU
ms.assetid: 425fbf08-d44c-2631-7400-55620fd429ee
ms.date: 06/08/2017
---


# Shape.CellsRowIndexU Property (Visio)

Returns the index of a row to which a cell belongs. Read-only.


## Syntax

 _expression_ . **CellsRowIndexU**( **_localeIndependentCellName_** )

 _expression_ An expression that returns a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _localeIndependentCellName_|Required| **String**|The universal name of a ShapeSheet cell.|

### Return Value

Integer


## Remarks

Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **CellsRowIndex** property to get a cell's row index by using the cell's local name. Use the **CellsRowIndexU** property to get a cell's row index by using the cell's universal name.


