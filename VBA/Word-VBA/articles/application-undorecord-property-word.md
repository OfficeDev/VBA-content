---
title: Application.UndoRecord Property (Word)
keywords: vbawd10.chm158335462
f1_keywords:
- vbawd10.chm158335462
ms.prod: word
api_name:
- Word.Application.UndoRecord
ms.assetid: d21c7089-2cdc-3d04-1073-ada649f21576
ms.date: 06/08/2017
---


# Application.UndoRecord Property (Word)

Returns an [UndoRecord](undorecord-object-word.md) object that provides a custom entry point into the undo stack. Read-only.


## Syntax

 _expression_ . **UndoRecord**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

Use the  **UndoRecord** object to create and modify custom undo records in the Word undo stack.


## Example

The following code example instantiates an  **UndoRecord** object.


```vb
Dim objUndo As UndoRecord 
Set objUndo = Application.UndoRecord
```


## See also


#### Concepts


[Application Object](application-object-word.md)

