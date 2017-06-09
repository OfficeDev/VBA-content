---
title: Image.Enabled Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: face613c-7a9c-9b28-ff79-656b83cbdf61
ms.date: 06/08/2017
---


# Image.Enabled Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether a control can receive the focus and respond to user-generated events. Read/write.


## Syntax

 _expression_. **Enabled**

 _expression_A variable that represents an  **Image** object.


## Remarks

 **True** is the control can receive the focus and respond to user-generated events, and is accessible through code (default). **False** if the user cannot interact with the control by using the mouse, keystrokes, accelerators, or hotkeys. The control is generally still accessible through code.

Use the  **Enabled** property to enable and disable controls. A disabled control appears dimmed, while an enabled control does not. Also, if a control displays a bitmap, the bitmap is dimmed whenever the control is dimmed. If **Enabled** is **False** for an **[Image](image-object-outlook-forms-script.md)**, the control does not initiate events but does not appear dimmed.


