---
title: Frame.CanPaste Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 520b845a-289f-9ed0-5af1-b5435462e027
ms.date: 06/08/2017
---


# Frame.CanPaste Property (Outlook Forms Script)

Returns a  **Boolean** that specifies whether the Clipboard contains data that the object supports. Read-only.


## Syntax

 _expression_. **CanPaste**

 _expression_A variable that represents a  **Frame** object.


## Remarks

 **True** if the object can receive information pasted from the Clipboard, **False** if the object cannot receive information pasted from the Clipboard.

 **CanPaste** is read-only.

If the Clipboard data is in a format that the object does not support, the  **CanPaste** property is **False**. For example, if you try to paste a bitmap into an object that only supports text,  **CanPaste** will be **False**.


