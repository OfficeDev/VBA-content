---
title: Shape.CellsRowIndex Property (Visio)
keywords: vis_sdr.chm11213200
f1_keywords:
- vis_sdr.chm11213200
ms.prod: visio
api_name:
- Visio.Shape.CellsRowIndex
ms.assetid: 7415afcb-9d98-5981-bd33-6ca18116470e
ms.date: 06/08/2017
---


# Shape.CellsRowIndex Property (Visio)

Returns the index of a row to which a cell belongs. Read-only.


## Syntax

 _expression_ . **CellsRowIndex**( **_localeSpecificCellName_** )

 _expression_ An expression that returns a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _localeSpecificCellName_|Required| **String**|The name of a ShapeSheet cell.|

### Return Value

Integer


## Remarks

The  **CellRowIndex** property searches by local cell names first; if it doesn't find the name, it will search by universal cell name.

Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **CellsRowIndex** property to get a cell's row index by using the cell's local name. Use the **CellsRowIndexU** property to get a cell's row index by using the cell's universal name.


