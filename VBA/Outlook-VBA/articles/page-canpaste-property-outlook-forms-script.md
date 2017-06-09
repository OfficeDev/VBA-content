---
title: Page.CanPaste Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: a9c14f8c-56e4-cebf-c423-c7bcc9a84cb6
ms.date: 06/08/2017
---


# Page.CanPaste Property (Outlook Forms Script)

Returns a  **Boolean** that specifies whether the Clipboard contains data that the object supports. Read-only.


## Syntax

 _expression_. **CanPaste**

 _expression_A variable that represents a  **Page** object.


## Remarks

 **True** if the object can receive information pasted from the Clipboard, **False** if the object cannot receive information pasted from the Clipboard.

 **CanPaste** is read-only.

If the Clipboard data is in a format that the object does not support, the  **CanPaste** property is **False**. For example, if you try to paste a bitmap into an object that only supports text,  **CanPaste** will be **False**.


