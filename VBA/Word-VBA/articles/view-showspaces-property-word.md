---
title: View.ShowSpaces Property (Word)
keywords: vbawd10.chm161808400
f1_keywords:
- vbawd10.chm161808400
ms.prod: word
api_name:
- Word.View.ShowSpaces
ms.assetid: c560747d-691a-1ddb-b748-2c91f519ba53
ms.date: 06/08/2017
---


# View.ShowSpaces Property (Word)

 **True** if space characters are displayed. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowSpaces**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Example

This example inserts spaces before the selection and displays space characters in the active window.


```vb
Selection.InsertBefore " " 
ActiveDocument.ActiveWindow.View.ShowSpaces = True
```


## See also


#### Concepts


[View Object](view-object-word.md)

