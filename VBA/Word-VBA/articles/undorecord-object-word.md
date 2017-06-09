---
title: UndoRecord Object (Word)
keywords: vbawd10.chm856
f1_keywords:
- vbawd10.chm856
ms.prod: word
api_name:
- Word.UndoRecord
ms.assetid: 77bf9801-e940-e661-6bbe-20a8714d5dbd
ms.date: 06/08/2017
---


# UndoRecord Object (Word)

Provides an entry point into the undo stack.


## Remarks

Use the  **UndoRecord** object to create and modify custom undo records in the Word undo stack.


## Example

The following code example instantiates an  **UndoRecord** object.


```vb
Dim objUndo As UndoRecord 
Set objUndo = Application.UndoRecord
```


## See also


#### Other resources


[Working With the UndoRecord Object](http://msdn.microsoft.com/library/e9df1047-5a1a-91da-3673-7e64b668552d%28Office.15%29.aspx)
[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


