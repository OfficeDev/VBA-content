---
title: Things you can do with a multicolumn ListBox or ComboBox
keywords: fm20.chm5225252
f1_keywords:
- fm20.chm5225252
ms.prod: office
ms.assetid: 99a32411-4a80-043c-b312-42fb3c3eb83f
ms.date: 06/08/2017
---


# Things you can do with a multicolumn ListBox or ComboBox

To control the column widths of a multicolumn  **ListBox** or **ComboBox**, you can specify the width, in points, for all the columns in the **ColumnWidths** property. Specifying zero for a specific column hides that column of information from the display.

If you want to hide all but one column of a  **ListBox** or **ComboBox** from the user, you can set the **ColumnWidths** of the other columns to zero and identify the column of information to display by leaving its **ColumnWidths** property set to the default value and using the **TextColumn** property. When the user selects a row, the **Text** property of the control is set to the value of the column identified by the **TextColumn** property. In a combo box, the system displays the column designated by the **TextColumn** property in the text box portion of the control.

Similarly, you can control which column of values is used for the control when the user makes a selection by specifying the column number in the  **BoundColumn** property.


