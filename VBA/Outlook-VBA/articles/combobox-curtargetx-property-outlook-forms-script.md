---
title: ComboBox.CurTargetX Property (Outlook Forms Script)
keywords: olfm10.chm2001020
f1_keywords:
- olfm10.chm2001020
ms.prod: outlook
ms.assetid: a12c1ba9-eca1-4a3f-89e4-1559b5e4b00c
ms.date: 06/08/2017
---


# ComboBox.CurTargetX Property (Outlook Forms Script)

Returns a  **Long** that represents the preferred horizontal position of the insertion point in a multiline **[ComboBox](combobox-object-outlook-forms-script.md)**. Read-only.


## Syntax

 _expression_. **CurTargetX**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

The  **CurTargetX** property retrieves the preferred position, measured in himetric units. A himetric is 0.0001 meter.

The target position is relative to the left edge of the control. If the length of a line is less than the value of the  **CurTargetX** property, you can place the insertion point at the end of the line. The value of **CurTargetX** changes when the user sets the insertion point or when the **[CurX](combobox-curx-property-outlook-forms-script.md)** property is set. **CurTargetX** is read-only.

The return value is valid when the object has focus.

You can use  **CurTargetX** and **CurX** to move the insertion point as the user scrolls through the contents of a multiline **ComboBox**. When the user moves the insertion point to another line of text by scrolling the content of the object,  **CurTargetX** specifies the preferred position for the insertion point. **CurX** is set to this value if the line of text is longer than the value of **CurTargetX**. Otherwise,  **CurX** is set to the end of the line of text.


