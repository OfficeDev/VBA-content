---
title: ListBox.Selected Property (Outlook Forms Script)
keywords: olfm10.chm2001830
f1_keywords:
- olfm10.chm2001830
ms.prod: outlook
ms.assetid: 653a977d-5ef8-0bd8-d851-927f03942a2c
ms.date: 06/08/2017
---


# ListBox.Selected Property (Outlook Forms Script)

Returns or sets a  **Boolean** that indicates the selection state of items in a **[ListBox](listbox-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **Selected**( **_pvargIndex_**)

 _expression_A variable that represents a  **ListBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|pvargIndex|Required| **Variant**|An integer with a range from 0 to one less than the number of items in the list.|

## Remarks

 **True** to indicate that the specified item is selected, **False** if it is not selected.

The  **Selected** property is useful when users can make multiple selections. You can use this property to determine the selected rows in a multi-select list box. You can also use this property to select or deselect rows in a list from code.

The default value of this property is based on the current selection state of the  **ListBox**.

For single-selection list boxes, the  **[Value](listbox-value-property-outlook-forms-script.md)** or **[ListIndex](listbox-listindex-property-outlook-forms-script.md)** properties are recommended for getting and setting the selection. In this case, **ListIndex** returns the index of the selected item. However, in a multiple selection, **ListIndex** returns the index of the row contained within the focus rectangle, regardless of whether the row is actually selected.

When a list box control's  **[MultiSelect](listbox-multiselect-property-outlook-forms-script.md)** property is set to 0, only one row can have its **Selected** property set to **True**.

Entering a value that is out of range for the index does not generate an error message, but does not set a property for any item in the list.


