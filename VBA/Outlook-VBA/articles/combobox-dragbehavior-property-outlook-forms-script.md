---
title: ComboBox.DragBehavior Property (Outlook Forms Script)
keywords: olfm10.chm2001085
f1_keywords:
- olfm10.chm2001085
ms.prod: outlook
ms.assetid: 38571166-8173-8612-54bd-f638044c2afb
ms.date: 06/08/2017
---


# ComboBox.DragBehavior Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies whether the system enables the drag-and-drop feature for the control. Read/write.


## Syntax

 _expression_. **DragBehavior**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

The possible values of  **DragBehavior** are 0 and 1. 0 represents that drag-and-drop action is not allowed. 1 represents that drag-and-drop action is allowed.

If the  **DragBehavior** property is enabled, dragging in a combo box starts a drag-and-drop operation on the selected text. If **DragBehavior** is disabled, dragging in a combo box selects text.

The drop-down portion of a  **[ComboBox](combobox-object-outlook-forms-script.md)** does not support drag-and-drop processes, nor does it support selection of list items within the text.

 **DragBehavior** has no effect on a **ComboBox** whose **[Style](combobox-style-property-outlook-forms-script.md)** property is set to 2.

You can combine the effects of the  **[EnterFieldBehavior](combobox-enterfieldbehavior-property-outlook-forms-script.md)** property and **DragBehavior** to create a large number of combo box styles.


