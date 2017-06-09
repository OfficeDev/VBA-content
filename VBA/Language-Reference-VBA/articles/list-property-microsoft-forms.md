---
title: List Property (Microsoft Forms)
keywords: fm20.chm2001400
f1_keywords:
- fm20.chm2001400
ms.prod: office
ms.assetid: 15ea715a-a361-34f4-98af-520942a6664e
ms.date: 06/08/2017
---


# List Property (Microsoft Forms)



Returns or sets the list entries of a  **ListBox** or **ComboBox**.
 **Syntax**
 _object_. **List(**_row, column_**)** [= _Variant_ ]
The  **List** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _row_|Required. An integer with a range from 0 to one less than the number of entries in the list.|
| _column_|Required. An integer with a range from 0 to one less than the number of columns.|
| _Variant_|Optional. The contents of the specified entry in the  **ListBox** or **ComboBox**.|
 **Settings**
Row and column numbering begins with zero. That is, the row number of the first row in the list is zero; the column number of the first column is zero. The number of the second row or column is 1, and so on.
 **Remarks**
The  **List** property works with the **ListCount** and **ListIndex** properties.Use **List** to access list items. A list is a variant[array](vbe-glossary.md); each item in the list has a row number and a column number.
Initially,  **ComboBox** and **ListBox** contain an empty list.

 **Note**  To specify items you want to display in a  **ComboBox** or **ListBox**, use the **AddItem** method. To remove items, use the **RemoveItem** method.

Use  **List** to copy an entire two-dimensional array of values to a control. Use **AddItem** to load a one-dimensional array or to load an individual element.

