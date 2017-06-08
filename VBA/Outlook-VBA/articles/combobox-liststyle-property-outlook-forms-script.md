---
title: ComboBox.ListStyle Property (Outlook Forms Script)
keywords: olfm10.chm2001450
f1_keywords:
- olfm10.chm2001450
ms.prod: outlook
ms.assetid: 9a061fe5-4c59-d051-97a1-db946a8ad8d4
ms.date: 06/08/2017
---


# ComboBox.ListStyle Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the visual appearance of the list in a **[ComboBox](combobox-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **ListStyle**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

The settings for  **ListStyle** are:



|**Value**|**Description**|
|:-----|:-----|
|0|Looks like a regular combo box, with the background of items highlighted.|
|1|Shows option buttons, or check boxes for a multi-select list of the combo box (default). When the user selects an item from the group, the option button associated with that item is selected and the option buttons for the other items in the group are deselected.|
The  **ListStyle** property lets you change the visual presentation of a **ComboBox**. By specifying a setting other than 0, you can present the contents of either control as a group of individual items, with each item including a visual cue to indicate whether it is selected.


