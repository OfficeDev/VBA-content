---
title: ComboBox.AutoSize Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 352cc408-0bef-43ae-c35b-38bea170507f
ms.date: 06/08/2017
---


# ComboBox.AutoSize Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether an object automatically resizes to display its entire contents. Read/write.


## Syntax

 _expression_. **AutoSize**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

 **True** to automatically resize the control to display its entire contents. **False** to keep the size of the control constant; contents are clipped when they exceed the area of the control (default).

For controls without captions, this property specifies whether the control automatically adjusts to display the information stored in the control. In a  **[ComboBox](combobox-object-outlook-forms-script.md)**, for example, setting  **AutoSize** to **True** automatically sets the width of the display area to match the length of the current text.

If you manually change the size of a control while  **AutoSize** is **True**, the manual change overrides the size previously set by  **AutoSize**.


