---
title: ListBox.List Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 3eb66479-c7d2-13d7-ebd3-1a09eb136dbe
ms.date: 06/08/2017
---


# ListBox.List Property (Outlook Forms Script)

Returns or sets a  **Variant** that represents the specified entry in a **[ListBox](listbox-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **List**( **_pvargIndex_**,  **_pvargColumn_**)

 _expression_A variable that represents a  **ListBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|pvargIndex|Optional| **Variant**|An integer with a range from 0 to one less than the number of entries in the list.|
|pvargColumn|Optional| **Variant**|An integer with a range from 0 to one less than the number of columns in the list.|

## Remarks

Row and column numbering begins with zero. That is, the row number of the first row in the list is zero; the column number of the first column is zero. The number of the second row or column is 1, and so on.

The  **List** property works with the **[ListCount](listbox-listcount-property-outlook-forms-script.md)** and **[ListIndex](listbox-listindex-property-outlook-forms-script.md)** properties. Use **List** to access list items. A list is a variant array; each item in the list has a row number and a column number.

Initially, a  **ListBox** contains an empty list.

To specify items you want to display in a  **ListBox**, use the  **[AddItem](listbox-additem-method-outlook-forms-script.md)** method. To remove items, use the **[RemoveItem](listbox-removeitem-method-outlook-forms-script.md)** method.

Use  **List** to copy an entire two-dimensional array of values to a control. Use **AddItem** to load a one-dimensional array or to load an individual element.


