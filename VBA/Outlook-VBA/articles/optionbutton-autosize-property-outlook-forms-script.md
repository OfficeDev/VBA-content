---
title: OptionButton.AutoSize Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 14f2dbc7-09e1-40a6-317d-b920c9290815
ms.date: 06/08/2017
---


# OptionButton.AutoSize Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether an object automatically resizes to display its entire contents. Read/write.


## Syntax

 _expression_. **AutoSize**

 _expression_A variable that represents an  **OptionButton** object.


## Remarks

 **True** to automatically resize the control to display its entire contents. **False** to keep the size of the control constant; contents are clipped when they exceed the area of the control (default).

For controls with captions, the  **AutoSize** property specifies whether the control automatically adjusts to display the entire caption.

If you manually change the size of a control while  **AutoSize** is **True**, the manual change overrides the size previously set by  **AutoSize**.


