---
title: ComboBox.ListIndex Property (Outlook Forms Script)
keywords: olfm10.chm2001430
f1_keywords:
- olfm10.chm2001430
ms.prod: outlook
ms.assetid: 2c4e473b-15e1-dce2-8748-30953b00a60f
ms.date: 06/08/2017
---


# ComboBox.ListIndex Property (Outlook Forms Script)

Returns or sets a  **Variant** that represents the currently selected item in a **[ComboBox](combobox-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **ListIndex**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

The  **ListIndex** property contains an index of the selected row in a list. Values of **ListIndex** range from -1 to one less than the total number of rows in a list (that is, ** [ListCount](combobox-listcount-property-outlook-forms-script.md)** - 1). When no rows are selected, **ListIndex** returns -1. When the user selects a row in a **ListBox** or **ComboBox**, the system sets the  **ListIndex** value. The **ListIndex** value of the first row in a list is 0, the value of the second row is 1, and so on.

The  **ListIndex** value is also available by setting the **[BoundColumn](combobox-boundcolumn-property-outlook-forms-script.md)** property to 0 for a combo box. If **BoundColumn** is 0, the underlying data source to which the combo box is bound contains the same list index value as **ListIndex**.


