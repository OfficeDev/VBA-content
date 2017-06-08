---
title: Page.UndoAction Method (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 2a8ff967-0f29-d986-312c-82cbd815b7e7
ms.date: 06/08/2017
---


# Page.UndoAction Method (Outlook Forms Script)

Reverses the most recent action that supports the  **Undo** command.


## Syntax

 _expression_. **UndoAction**

 _expression_A variable that represents a  **Page** object.


## Remarks

Not all user actions can be undone. If an action cannot be undone, the  **Undo** command is unavailable following the action.

You must apply this method before the form or control is updated. You may want to include this method in a form's  **PropertyChange** event.


