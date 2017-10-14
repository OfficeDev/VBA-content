---
title: ComboBox.ColumnWidths Property (Outlook Forms Script)
keywords: olfm10.chm2000960
f1_keywords:
- olfm10.chm2000960
ms.prod: outlook
ms.assetid: 434b2e52-fe2c-78ff-2fc6-1c0b9b858ebc
ms.date: 06/08/2017
---


# ComboBox.ColumnWidths Property (Outlook Forms Script)

Returns or sets a  **String** that specifies the width of each column in a multicolumn **[ComboBox](combobox-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **ColumnWidths**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

 **ColumnWidths** sets the column width in points. A setting of -1 or blank results in a calculated width. A width of 0 hides a column. To specify a different unit of measurement, include the unit of measure. A value greater than 0 explicitly specifies the width of the column.

To separate column entries, use semicolons (;) as list separators. Or use the list separator specified in  **Regional Settings** of the Windows Control Panel.

Any or all of the  **ColumnWidths** property settings can be blank. You create a blank setting by typing a list separator without a preceding value.

If you specify a -1 in the property page, the displayed value in the property page is a blank.

To calculate column widths when  **ColumnWidths** is blank or -1, the width of the control is divided equally among all columns of the list. If the sum of the specified column widths exceeds the width of the control, the list is left-aligned within the control and one or more of the rightmost columns are not displayed. Users can scroll the list using the horizontal scroll bar to display the rightmost columns.

The minimum calculated column width is 72 points (1 inch). To produce columns narrower than this, you must specify the width explicitly.

Unless specified otherwise, column widths are measured in points. To specify another unit of measure, include the units as part of the values. The following examples specify column widths in several units of measure and describe how the various settings would fit in a three-column list box that is 4 inches wide.



|**Setting**|**Effect**|
|:-----|:-----|
|90;72;90|The first column is 90 points (1.25 inch); the second column is 72 points (1 inch); the third column is 90 points.|
|6 cm;0;6 cm|The first column is 6 centimeters; the second column is hidden; the third column is 6 centimeters. Because part of the third column is visible, a horizontal scroll bar appears.|
|1.5 in;0;2.5 in|The first column is 1.5 inches, the second column is hidden, and the third column is 2.5 inches.|
|2 in;;2 in|The first column is 2 inches, the second column is 1 inch (default), and the third column is 2 inches. Because only half of the third column is visible, a horizontal scroll bar appears.|
|(Blank)|All three columns are the same width (1.33 inches).|
In a  **ComboBox**, the system displays the column designated by the  **[TextColumn](combobox-textcolumn-property-outlook-forms-script.md)** property in the text box portion of the control. Setting **TextColumn** to -1 displays the first column that has a **ColumnWidths** value greater than 0.


