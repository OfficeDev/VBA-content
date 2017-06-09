---
title: ListBox.AddItem Method (Outlook Forms Script)
ms.prod: outlook
ms.assetid: e948d5ac-6d88-d825-e1ee-4a05fe934853
ms.date: 06/08/2017
---


# ListBox.AddItem Method (Outlook Forms Script)

For a single-column  **[ListBox](listbox-object-outlook-forms-script.md)**, the  **AddItem** method adds an item to the list. For a multicolumn **ListBox**, this method adds a row to the list.


## Syntax

 _expression_. **AddItem**( **_pvargItem_**,  **_pvargIndex_**)

 _expression_A variable that represents a  **ListBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|pvargItem|Optional| **Variant**|Specifies the item or row to add. The number of the first item or row is 0; the number of the second item or row is 1, and so on.|
|pvargIndex|Optional| **Variant**|Integer specifying the position within the object where the new item or row is placed.|

## Remarks

If you supply a valid value for  _varIndex_, the  **AddItem** method places the item or row at that position within the list. If you omit _varIndex_, the method adds the item or row at the end of the list.

The value of  _varIndex_ must not be greater than the value of the **[ListCount](listbox-listcount-property-outlook-forms-script.md)** property.

For a multicolumn  **ListBox**,  **AddItem** inserts an entire row, that is, it inserts an item for each column of the control. To assign values to an item beyond the first column, use the **[List](listbox-list-property-outlook-forms-script.md)** or **[Column](listbox-column-property-outlook-forms-script.md)** property and specify the row and column of the item.

If the control is bound to data, the  **AddItem** method fails.

You can add more than one row at a time to a  **ListBox** by using **List**.


