---
title: Frame.UndoAction Method (Outlook Forms Script)
keywords: olfm10.chm2000450
f1_keywords:
- olfm10.chm2000450
ms.prod: outlook
ms.assetid: 28ca1383-bfd1-db6c-2945-82dd29a3b9ae
ms.date: 06/08/2017
---


# Frame.UndoAction Method (Outlook Forms Script)

Reverses the most recent action that supports the  **Undo** command.


## Syntax

 _expression_. **UndoAction**

 _expression_A variable that represents a  **Frame** object.


### Return Value

A Boolean that is  **True** if the method succeeds, **False** otherwise.


## Remarks

Not all user actions can be undone. If an action cannot be undone, the  **Undo** command is unavailable following the action.

You must apply this method before the form or control is updated. You may want to include this method in a form's  **PropertyChange** event.


