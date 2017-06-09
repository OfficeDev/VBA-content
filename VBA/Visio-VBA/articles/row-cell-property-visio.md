---
title: Row.Cell Property (Visio)
keywords: vis_sdr.chm15813180
f1_keywords:
- vis_sdr.chm15813180
ms.prod: visio
api_name:
- Visio.Row.Cell
ms.assetid: 74613af7-4c01-aa91-3659-22e313cd6d2c
ms.date: 06/08/2017
---


# Row.Cell Property (Visio)

Uses the local name or index of a cell to return the cell. Read-only.


## Syntax

 _expression_ . **Cell**( **_NameOrIndex_** )

 _expression_ A variable that represents a **Row** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NameOrIndex_|Required| **Variant**|The local name or index of the cell.|

### Return Value

Cell


## Remarks

The first cell in a row has an index of zero (0). 




 **Note**  Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

 As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the **Cell** property to get a **Cell** object by using its local name. Use the **CellU** property to get a **Cell** object by using its universal name.


