---
title: TabStrip.Enabled Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: ab76f52f-5b9c-b8f9-b1b8-f29ff495d5e1
ms.date: 06/08/2017
---


# TabStrip.Enabled Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether a control can receive the focus and respond to user-generated events. Read/write.


## Syntax

 _expression_. **Enabled**

 _expression_A variable that represents a  **TabStrip** object.


## Remarks

 **True** is the control can receive the focus and respond to user-generated events, and is accessible through code (default). **False** if the user cannot interact with the control by using the mouse, keystrokes, accelerators, or hotkeys. The control is generally still accessible through code.

Use the  **Enabled** property to enable and disable controls. A disabled control appears dimmed, while an enabled control does not. Also, if a control displays a bitmap, the bitmap is dimmed whenever the control is dimmed.

When the user tabs into an enabled  **[TabStrip](tabstrip-object-outlook-forms-script.md)**, the first tab in the control receives the focus. If the first tab of a  **TabStrip** is disabled, the first enabled tab of that control receives the focus. If all tabs of a **TabStrip** are disabled, the control is disabled and cannot receive the focus.


