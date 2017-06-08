---
title: View.ShowFormatChanges Property (Word)
keywords: vbawd10.chm161808421
f1_keywords:
- vbawd10.chm161808421
ms.prod: word
api_name:
- Word.View.ShowFormatChanges
ms.assetid: e431dc24-a975-958c-07dc-64062e05cb26
ms.date: 06/08/2017
---


# View.ShowFormatChanges Property (Word)

 **True** for Microsoft Word to display formatting changes made to a document with Track Changes enabled. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowFormatChanges**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Example

This example hides the formatting changes made to the active document. This example assumes that formatting changes have been made to a document in which Track Changes is enabled.


```vb
Sub HideFormattingChanges() 
 ActiveWindow.View.ShowFormatChanges = False 
End Sub
```


## See also


#### Concepts


[View Object](view-object-word.md)

