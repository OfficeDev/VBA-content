---
title: RedoAction Method
keywords: fm20.chm5224966
f1_keywords:
- fm20.chm5224966
ms.prod: office
api_name:
- Office.RedoAction
ms.assetid: a4aba525-5cbe-1a68-aec6-731fb5f78464
ms.date: 06/08/2017
---


# RedoAction Method



Reverses the effect of the most recent Undo action.
 **Syntax**
 _Boolean_ = _object_. **RedoAction**
The  **RedoAction** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
 **Remarks**
Redo reverses the last Undo, which is not necessarily the last action taken. Not all actions can be undone.
For example, after pasting text into a  **TextBox** and then choosing the Undo command to remove the text, you can choose the Redo command to put the text back in.

 **Note**  If the  **CanRedo** property is **False**, the Redo command is not available in the user interface, and the **RedoAction** method is not valid in code.

 **RedoAction** returns **True** if it was successful.

