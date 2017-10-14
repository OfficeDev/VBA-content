---
title: CommandButton.AutoSize Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 85bc6640-679a-955a-1bc0-9bab0691fae8
ms.date: 06/08/2017
---


# CommandButton.AutoSize Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether an object automatically resizes to display its entire contents. Read/write.


## Syntax

 _expression_. **AutoSize**

 _expression_A variable that represents a  **CommandButton** object.


## Remarks

 **True** to automatically resize the control to display its entire contents. **False** to keep the size of the control constant; contents are clipped when they exceed the area of the control (default).

For controls with captions, the  **AutoSize** property specifies whether the control automatically adjusts to display the entire caption.

If you manually change the size of a control while  **AutoSize** is **True**, the manual change overrides the size previously set by  **AutoSize**.


