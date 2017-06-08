---
title: Page.Enabled Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: ecfd8efd-d86d-a483-7869-016cb16cc37d
ms.date: 06/08/2017
---


# Page.Enabled Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether a control can receive the focus and respond to user-generated events. Read/write.


## Syntax

 _expression_. **Enabled**

 _expression_A variable that represents a  **Page** object.


## Remarks

 **True** is the control can receive the focus and respond to user-generated events, and is accessible through code (default). **False** if the user cannot interact with the control by using the mouse, keystrokes, accelerators, or hotkeys. The control is generally still accessible through code.

Use the  **Enabled** property to enable and disable controls. A disabled control appears dimmed, while an enabled control does not. Also, if a control displays a bitmap, the bitmap is dimmed whenever the control is dimmed.

When the user tabs into an enabled  **[MultiPage](multipage-object-outlook-forms-script.md)**, the first page in the control receives the focus. If the first page of a  **MultiPage** is disabled, the first enabled page of that control receives the focus. If all pages of a **MultiPage** are disabled, the control is disabled and cannot receive the focus.


