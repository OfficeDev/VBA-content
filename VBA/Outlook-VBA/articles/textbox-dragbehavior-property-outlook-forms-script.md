---
title: TextBox.DragBehavior Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 625ff366-65d5-0b50-bd73-420df5324fd2
ms.date: 06/08/2017
---


# TextBox.DragBehavior Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies whether the system enables the drag-and-drop feature for the control. Read/write.


## Syntax

 _expression_. **DragBehavior**

 _expression_A variable that represents a  **TextBox** object.


## Remarks

The possible values of  **DragBehavior** are 0 and 1. 0 represents that drag-and-drop action is not allowed. 1 represents that drag-and-drop action is allowed.

If the  **DragBehavior** property is enabled, dragging in a text box starts a drag-and-drop operation on the selected text. If **DragBehavior** is disabled, dragging in a text box selects text.

You can combine the effects of the  **[EnterFieldBehavior](textbox-enterfieldbehavior-property-outlook-forms-script.md)** property and **DragBehavior** to create a large number of text box styles.


