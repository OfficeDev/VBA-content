---
title: View.ShowDrawings Property (Word)
keywords: vbawd10.chm161808398
f1_keywords:
- vbawd10.chm161808398
ms.prod: word
api_name:
- Word.View.ShowDrawings
ms.assetid: fa03b2f0-e090-5130-c370-4a00ee6db958
ms.date: 06/08/2017
---


# View.ShowDrawings Property (Word)

 **True** if objects created with the drawing tools are displayed in print layout view. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowDrawings**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Example

This example switches the active window to print layout view and displays objects created with the drawing tools.


```vb
With ActiveDocument.ActiveWindow.View 
 .Type = wdPrintView 
 .ShowDrawings = True 
End With
```


## See also


#### Concepts


[View Object](view-object-word.md)

