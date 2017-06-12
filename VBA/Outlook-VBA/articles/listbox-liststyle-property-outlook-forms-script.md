---
title: ListBox.ListStyle Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 4abbd557-b80f-e940-873f-8527e30b4a2e
ms.date: 06/08/2017
---


# ListBox.ListStyle Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the visual appearance of the list in a **[ListBox](listbox-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **ListStyle**

 _expression_A variable that represents a  **ListBox** object.


## Remarks

The settings for  **ListStyle** are:



|**Value**|**Description**|
|:-----|:-----|
|0|Looks like a regular list box, with the background of items highlighted.|
|1|Shows option buttons, or check boxes for a multi-select list (default). When the user selects an item from the group, the option button associated with that item is selected and the option buttons for the other items in the group are deselected.|
The  **ListStyle** property lets you change the visual presentation of a **ListBox**. By specifying a setting other than 0, you can present the contents of either control as a group of individual items, with each item including a visual cue to indicate whether it is selected.

If the list box supports a single selection (the  **[MultiSelect](listbox-multiselect-property-outlook-forms-script.md)** property is set to 0), the user can press one button in the group. If the control supports multi-select, the user can press two or more buttons in the group.


