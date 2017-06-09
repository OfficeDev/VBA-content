---
title: ListBox.ColumnCount Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 8ae3ba58-4ac6-4609-b159-2b353037b949
ms.date: 06/08/2017
---


# ListBox.ColumnCount Property (Outlook Forms Script)

Returns or sets a  **Long** that represents the number of columns to display in a list box. Read/write.


## Syntax

 _expression_. **ColumnCount**

 _expression_A variable that represents a  **ListBox** object.


## Remarks

If you set the  **ColumnCount** property for a list box to 3 on an employee form, one column can list last names, another can list first names, and the third can list employee ID numbers.

Setting  **ColumnCount** to 0 displays zero columns, and setting it to -1 displays all the available columns. For an unbound data source, there is a 10-column limit (0 to 9).

You can use the  **[ColumnWidths](listbox-columnwidths-property-outlook-forms-script.md)** property to set the width of the columns displayed in the control.


