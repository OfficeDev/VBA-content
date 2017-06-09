---
title: ColumnCount Property
keywords: fm20.chm2000940
f1_keywords:
- fm20.chm2000940
ms.prod: office
api_name:
- Office.ColumnCount
ms.assetid: ba998cac-3e31-eb81-8f35-fe7fee133e63
ms.date: 06/08/2017
---


# ColumnCount Property



Specifies the number of columns to display in a list box or combo box.
 **Syntax**
 _object_. **ColumnCount** [= _Long_ ]
The  **ColumnCount** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Long_|Optional. Specifies the number of columns to display.|
 **Remarks**
If you set the  **ColumnCount** property for a list box to 3 on an employee form, one column can list last names, another can list first names, and the third can list employee ID numbers.
Setting  **ColumnCount** to 0 displays zero columns, and setting it to -1 displays all the available columns. For an[unbound](glossary-vba.md)[data source](glossary-vba.md), there is a 10-column limit (0 to 9).
You can use the  **ColumnWidths** property to set the width of the columns displayed in the control.

