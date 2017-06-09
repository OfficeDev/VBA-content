---
title: View.ShowTextBoundaries Property (Word)
keywords: vbawd10.chm161808396
f1_keywords:
- vbawd10.chm161808396
ms.prod: word
api_name:
- Word.View.ShowTextBoundaries
ms.assetid: a9bc7cc0-0062-0b1d-6e16-19ed52ba9fb9
ms.date: 06/08/2017
---


# View.ShowTextBoundaries Property (Word)

 **True** if dotted lines are displayed around page margins, text columns, objects, and frames in print layout view. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowTextBoundaries**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Example

This example switches the active window to page view and displays text boundary lines.


```vb
With ActiveDocument.ActiveWindow.View 
 .Type = wdPrintView 
 .ShowTextBoundaries = True 
End With
```


## See also


#### Concepts


[View Object](view-object-word.md)

