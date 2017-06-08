---
title: ComboBox.List Property (Outlook Forms Script)
keywords: olfm10.chm2001400
f1_keywords:
- olfm10.chm2001400
ms.prod: outlook
ms.assetid: 687f44e8-7b4b-eab5-93b8-022cd4d1c302
ms.date: 06/08/2017
---


# ComboBox.List Property (Outlook Forms Script)

Returns or sets a  **Variant** that represents the specified entry in a **[ComboBox](combobox-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **List**( **_pvargIndex_**,  **_pvargColumn_**)

 _expression_A variable that represents a  **ComboBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|pvargIndex|Optional| **Variant**|An integer with a range from 0 to one less than the number of entries in the list of the  **ComboBox**.|
|pvargColumn|Optional| **Variant**|An integer with a range from 0 to one less than the number of columns in the list of the  **ComboBox**.|

## Remarks

Row and column numbering begins with zero. That is, the row number of the first row in the list is zero; the column number of the first column is zero. The number of the second row or column is 1, and so on.

The  **List** property works with the **[ListCount](combobox-listcount-property-outlook-forms-script.md)** and **[ListIndex](combobox-listindex-property-outlook-forms-script.md)** properties. Use **List** to access list items. A list is a variant array; each item in the list has a row number and a column number.

Initially, a  **ComboBox** contains an empty list.

To specify items you want to display in a  **ComboBox**, use the  **[AddItem](combobox-additem-method-outlook-forms-script.md)** method. To remove items, use the **[RemoveItem](combobox-removeitem-method-outlook-forms-script.md)** method.

Use  **List** to copy an entire two-dimensional array of values to a control. Use **AddItem** to load a one-dimensional array or to load an individual element.


