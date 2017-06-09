---
title: TextBox.EnterFieldBehavior Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: b160b411-80b6-8731-3ee8-ac7ab889daf0
ms.date: 06/08/2017
---


# TextBox.EnterFieldBehavior Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the selection behavior when entering a **[TextBox](textbox-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **EnterFieldBehavior**

 _expression_A variable that represents a  **TextBox** object.


## Remarks

The possible values of  **EnterFieldBehavior** are 0 and 1. 0 represents selecting the entire contents of the edit region when entering the control (default). 1 represents leaving the selection unchanged. Visually, this uses the selection that was in effect the last time the control was active.

The  **EnterFieldBehavior** property controls the way text is selected when the user tabs to the control, not when the control receives focus as a result of the **SetFocus** method. Following **SetFocus**, the contents of the control are not selected and the insertion point appears after the last character in the control's edit region.

You can combine the effects of the  **EnterFieldBehavior** property and **[DragBehavior](textbox-dragbehavior-property-outlook-forms-script.md)** to create a large number of text box styles.


