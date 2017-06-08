---
title: Cell.RowName Property (Visio)
keywords: vis_sdr.chm10114260
f1_keywords:
- vis_sdr.chm10114260
ms.prod: visio
api_name:
- Visio.Cell.RowName
ms.assetid: 4f5f57f9-c147-5991-c3f0-2caad2993d77
ms.date: 06/08/2017
---


# Cell.RowName Property (Visio)

Gets or sets the name of the row that contains the  **Cell** object. Read/write.


## Syntax

 _expression_ . **RowName**

 _expression_ A variable that represents a **Cell** object.


### Return Value

String


## Remarks

If the cell is in a row in a shape's User-Defined Cells, Shape Data, or Connection Points section, the  **RowName** property can get or set the name of the row. If the cell is not in one of these sections, attempting to get or set the name generates an error.

The Connection Points section can contain either named or unnamed rows, but not a combination of the two. Getting the name of an unnamed Connection Points row returns a zero-length string ("") and does not generate an error. Setting the name of an unnamed row in a Connection Points row assigns the name to the target row and converts all remaining rows in the section to named rows, using their default names (Row_1, Row_2, and so on). Assigning a zero-length string ("") to a named row in a Connection Points section resets the named row to its default name, but has no effect on an unnamed Connection Points row.

When you change a row name, any cell objects referring to cells in that row become invalid and you must reassign them. Also, if other Connection Points rows become named as a result of a row name change, you must also reassign references to cells in those rows. 


 **Note**  Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **RowName** property to get or set an object's local row name. Use the **RowNameU** property to get or set an object's universal row name.


