---
title: UndoAction Method
keywords: fm20.chm5224975
f1_keywords:
- fm20.chm5224975
ms.prod: office
api_name:
- Office.UndoAction
ms.assetid: 751fb2c5-4fa6-bab5-fb9a-5c396d05cae1
ms.date: 06/08/2017
---


# UndoAction Method



Reverses the most recent action that supports the Undo command.
 **Syntax**
 _Boolean_ = _object_. **UndoAction**
The  **UndoAction** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
 **Remarks**
The Undo command in the user interface uses the  **UndoAction** method. For example, if you paste text into a **TextBox**, you can use **UndoAction** to remove that text and restore the previous contents of the **TextBox**.
Not all user actions can be undone. If an action cannot be undone, the Undo command is unavailable following the action.

 **Note**  If the  **CanUndo** property is **False**, the Undo command is not available in the user interface, and **UndoAction** is not valid in code.

If  **UndoAction** is applied to a form, all changes to the current record are lost. If **UndoAction** is applied to a control, only the control itself is affected.
You must apply this method before the form or control is updated. You may want to include this method in a form's BeforeUpdate event or a control's Change event.
 **UndoAction** is an alternative to using the[SendKeys Statement](vbe-glossary.md) to send the value of ESC in an event procedure.

