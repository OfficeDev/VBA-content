---
title: TextColumn Property
keywords: fm20.chm2002090
f1_keywords:
- fm20.chm2002090
ms.prod: office
api_name:
- Office.TextColumn
ms.assetid: 65a18466-3a31-d3a8-4585-eb0ba3a6e473
ms.date: 06/08/2017
---


# TextColumn Property



Identifies the column in a  **ComboBox** or **ListBox** to store in the **Text** property when the user selects a row.
 **Syntax**
 _object_. **TextColumn** [= _Variant_ ]
The  **TextColumn** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Variant_|Optional. The column to be displayed.|
 **Settings**
Values for the  **TextColumn** property range from -1 to the number of columns in the list. The **TextColumn** value for the first column is 1, the value of the second column is 2, and so on. Setting **TextColumn** to 0 displays the **ListIndex** values. Setting **TextColumn** to -1 displays the first column that has a **ColumnWidths** value greater than 0.
 **Remarks**
In a combo box, the system displays the column designated by the  **TextColumn** property in the text box portion of the control.
When the user selects a row from a  **ComboBox** or **ListBox**, the column referenced by **TextColumn** is stored in the **Text** property.For example, you could set up a multicolumn **ListBox** that contains the names of holidays in one column and dates for the holidays in a second column. To present the holiday names to users, specify the first column as the **TextColumn**. To store the dates of the holidays, specify the second column as the **BoundColumn**. To hide the dates of the holidays, set the **ColumnWidths** property of the sceond column to zero.
When the  **Text** property of a **ComboBox** changes (such as when a user types an entry into the control), the new text is compared to the column of data specified by **TextColumn**.

