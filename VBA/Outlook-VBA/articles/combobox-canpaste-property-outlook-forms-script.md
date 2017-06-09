---
title: ComboBox.CanPaste Property (Outlook Forms Script)
keywords: olfm10.chm2000850
f1_keywords:
- olfm10.chm2000850
ms.prod: outlook
ms.assetid: 36b1909a-fe23-77f9-4072-0264a6be02c8
ms.date: 06/08/2017
---


# ComboBox.CanPaste Property (Outlook Forms Script)

Returns a  **Boolean** that specifies whether the Clipboard contains data that the object supports. Read-only.


## Syntax

 _expression_. **CanPaste**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

 **True** if the object can receive information pasted from the Clipboard, **False** if the object cannot receive information pasted from the Clipboard.

 **CanPaste** is read-only.

If the Clipboard data is in a format that the object does not support, the  **CanPaste** property is **False**. For example, if you try to paste a bitmap into an object that only supports text,  **CanPaste** will be **False**.


