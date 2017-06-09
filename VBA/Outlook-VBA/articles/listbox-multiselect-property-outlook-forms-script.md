---
title: ListBox.MultiSelect Property (Outlook Forms Script)
keywords: olfm10.chm2001580
f1_keywords:
- olfm10.chm2001580
ms.prod: outlook
ms.assetid: 4ecc299b-0733-aa23-e820-f341ac80a0fa
ms.date: 06/08/2017
---


# ListBox.MultiSelect Property (Outlook Forms Script)

Returns or sets an  **Integer** that indicates whether the object permits multiple selections. Read/write.


## Syntax

 _expression_. **MultiSelect**

 _expression_A variable that represents a  **ListBox** object.


## Remarks

The settings for  **MultiSelect** are:



|**Value**|**Description**|
|:-----|:-----|
|0|Only one item can be selected (default).|
|1|Pressing the  **SPACEBAR** or clicking selects or deselects an item in the list.|
|2|Pressing  **SHIFT** and clicking the mouse, or pressing **SHIFT** and one of the arrow keys, extends the selection from the previously selected item to the current item. Pressing **CTRL** and clicking the mouse selects or deselects an item.|
When the  **MultiSelect** property is set to 1 or 2, you must use the list box's **[Selected](listbox-selected-property-outlook-forms-script.md)** property to determine the selected items. Also, the **[Value](listbox-value-property-outlook-forms-script.md)** property of the control is always **Null**.

The  **[ListIndex](listbox-listindex-property-outlook-forms-script.md)** property returns the index of the row with the keyboard focus.


