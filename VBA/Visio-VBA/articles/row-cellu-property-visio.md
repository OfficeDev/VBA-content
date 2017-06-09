---
title: Row.CellU Property (Visio)
keywords: vis_sdr.chm15851960
f1_keywords:
- vis_sdr.chm15851960
ms.prod: visio
api_name:
- Visio.Row.CellU
ms.assetid: 1fd467e1-9c5e-238a-b7d6-253668f94882
ms.date: 06/08/2017
---


# Row.CellU Property (Visio)

Uses the universal name or index of a cell to return the cell. Read-only.


## Syntax

 _expression_ . **CellU**( **_NameOrIndex_** )

 _expression_ A variable that represents a **Row** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NameOrIndex_|Required| **Variant**|The universal name or index of the cell.|

### Return Value

Cell


## Remarks

The first cell in a row has an index of zero (0).




 **Note**  Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **Cell** property to get a **Cell** object by using its local name. Use the **CellU** property to get a **Cell** object by using its universal name.


