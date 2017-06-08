---
title: ComboBox.EnterFieldBehavior Property (Outlook Forms Script)
keywords: olfm10.chm2001125
f1_keywords:
- olfm10.chm2001125
ms.prod: outlook
ms.assetid: dffb2409-fc12-7632-58e4-118f331072a7
ms.date: 06/08/2017
---


# ComboBox.EnterFieldBehavior Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the selection behavior when entering a **[ComboBox](combobox-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **EnterFieldBehavior**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

The possible values of  **EnterFieldBehavior** are 0 and 1. 0 represents selecting the entire contents of the edit region when entering the control (default). 1 represents leaving the selection unchanged. Visually, this uses the selection that was in effect the last time the control was active.

The  **EnterFieldBehavior** property controls the way text is selected when the user tabs to the control, not when the control receives focus as a result of the **SetFocus** method. Following **SetFocus**, the contents of the control are not selected and the insertion point appears after the last character in the control's edit region.

You can combine the effects of the  **EnterFieldBehavior** property and **[DragBehavior](olkcombobox-dragbehavior-property-outlook.md)** to create a large number of combo box styles.


