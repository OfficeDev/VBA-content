---
title: ListBox.ColumnWidths Property (Access)
keywords: vbaac10.chm11226
f1_keywords:
- vbaac10.chm11226
ms.prod: access
api_name:
- Access.ListBox.ColumnWidths
ms.assetid: 4ac2a001-8084-37aa-9f8e-ec3d373f7161
ms.date: 06/08/2017
---


# ListBox.ColumnWidths Property (Access)

You can use the  **ColumnWidths** property to specify the width of each column in a multiple-column list box. Read/write **String**.


## Syntax

 _expression_. **ColumnWidths**

 _expression_ A variable that represents a **ListBox** object.


## Remarks

The  **ColumnWidths** property holds a value specifying the width of each column in inches or centimeters, depending on the measurement system (U.S. or Metric) selected in the **Measurement system** box on the **Number** tab of the **Regional Options** dialog box of Windows Control Panel. The default setting is 1 inch or 2.54 centimeters. The **ColumnWidths** property setting must be a value from 0 to 22 inches (55.87 cm) for each column in the list box or combo box.

To separate your column entries, use semicolons (;) as list separators (or the list separator selected in the  **List separator** box on the **Number** tab of the **Regional Options** dialog box).

A width of 0 hides a column. Any or all of the  **ColumnWidths** property settings can be blank. You create a blank setting by typing a list separator without a preceding value. Blank values result in Microsoft Access automatically setting a default column width that varies depending on the number of columns and the width of the combo box or list box.

In Visual Basic, use a string expression to set the column width values in twips. Column widths are separated by semicolons. To specify a different unit of measurement, include the unit of measure (cm or in). For example, the following string expression specifies three column widths in centimeters.




```
"6 cm;0;6 cm"
```

You can also use this property to hide one or more columns.

If you leave the  **ColumnWidths** property setting blank, Microsoft Access sets the width of each column as the overall width of the list box or combo box divided by the number of columns.

If the column widths you set are too wide to be fully displayed within the combo box or list box, the rightmost columns are hidden and a horizontal scroll bar appears.

If you specify the width for some columns but leave the setting for others blank, Microsoft Access divides the remaining width by the number of columns for which you haven't specified a width. The minimum calculated column width is 1,440 twips (1 inch).

For example, the following settings are applied to a 4-inch list box with three columns.



|**Setting**|**Description**|
|:-----|:-----|
|1.5 in;0;2.5 in|The first column is 1.5 inches, the second column is hidden, and the third column is 2.5 inches.|
|2 in;;2 in|The first column is 2 inches, the second column is 1 inch (default), and the third column is 2 inches. Because only half of the third column is visible, a horizontal scroll bar appears.|
|(Blank)|The three columns are the same width (1.33 inches).|

 **Note**  This property is different than the  **ColumnWidth** property, which specifies the width of a specified column in a datasheet.


## See also


#### Concepts


[ListBox Object](listbox-object-access.md)

