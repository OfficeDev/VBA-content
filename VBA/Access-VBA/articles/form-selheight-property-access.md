---
title: Form.SelHeight Property (Access)
keywords: vbaac10.chm13472
f1_keywords:
- vbaac10.chm13472
ms.prod: access
api_name:
- Access.Form.SelHeight
ms.assetid: c8808132-ab4d-77f1-fbf7-121d37e188fe
ms.date: 06/08/2017
---


# Form.SelHeight Property (Access)

You can use the  **SelHeight** property to specify or determine the number of selected rows (records) in the current selection rectangle in a table, query, or form datasheet, or the number of selected records in a continuous form. Read/write **Long**.


## Syntax

 _expression_. **SelHeight**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **SelHeight** property returns a **Long Integer** value between 0 and the number of records in the datasheet or continuous form. The setting of this property specifies or returns the number of selected rows in the selection rectangle or the number of selected records in the continuous form.

If there's no selection, the value returned by this property will be zero. Setting this property to 0 removes the selection from the datasheet or form.

If you've selected one or more columns in a datasheet (using the column headings), you can't change the setting of the  **SelHeight** property (except to set it to 0).

You can use these properties with the  **SelTop** and **SelLeft** properties to specify or determine the actual position of the selection rectangle on the datasheet. If there's no selection, then the **SelTop** and **SelLeft** properties return the row number and column number of the cell with the focus.

The  **SelHeight** and **SelWidth** properties contain the position of the lower-right corner of the selection rectangle. The **SelTop** and **SelLeft** property values determine the upper-left corner of the selection rectangle.


## See also


#### Concepts


[Form Object](form-object-access.md)

