---
title: BoundColumn Property
keywords: fm20.chm5225011
f1_keywords:
- fm20.chm5225011
ms.prod: office
api_name:
- Office.BoundColumn
ms.assetid: 6c5c5c31-0bd3-87bf-4c1d-0b1064ffc0d6
ms.date: 06/08/2017
---


# BoundColumn Property



Identifies the source of data in a multicolumn  **ComboBox** or **ListBox**.
 **Syntax**
 _object_. **BoundColumn** [= _Variant_ ]
The  **BoundColumn** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Variant_|Optional. Indicates how the  **BoundColumn** value is selected.|
 **Settings**
The settings for  _Variant_ are:


|**Value**|**Description**|
|:-----|:-----|
|0|Assigns the value of the  **ListIndex** property to the control.|
|1 or greater|Assigns the value from the specified column to the control. Columns are numbered from 1 when using this property (default).|
 **Remarks**
When the user chooses a row in a multicolumn  **ListBox** or **ComboBox**, the **BoundColumn** property identifies which item from that row to store as the value of the control. For example, if each row contains 8 items and **BoundColumn** is 3, the system stores the information in the third column of the currently-selected row as the value of the object.
You can display one set of data to users but store different, associated values for the object by using the  **BoundColumn** and the **TextColumn** properties. **TextColumn** identifies the column of data displayed in text box portion of a **ComboBox** and the value stored in the **Text** property; **BoundColumn** identifies the column of associated data values stored for the control. For example, you could set up a multicolumn **ListBox** that contains the names of holidays in one column and dates for the holidays in a second column. To present the holiday names to users, specify the first column as the **TextColumn**. To store the dates of the holidays, specify the second column as the **BoundColumn**. To hide the dates of the holidays, set the **ColumnWidths** property of the sceond column to zero.
If the control is [bound](glossary-vba.md) to a[data source](glossary-vba.md), the value in the column specified by  **BoundColumn** is stored in the data source named in the **ControlSource** property.
The  **ListIndex** value retrieves the number of the selected row. For example, if you want to know the row of the selected item, set **BoundColumn** to 0 to assign the number of the selected row as the value of the control. Be sure to retrieve a current value, rather than relying on a previously saved value, if you are referencing a list whose contents might change.
The  **Column**, **List**, and **ListIndex** properties all use zero-based numbering. That is, the value of the first item (column or row) is zero; the value of the second item is one, and so on. This means that if **BoundColumn** is set to 3, you could access the value stored in that column using the expression Column(2).

