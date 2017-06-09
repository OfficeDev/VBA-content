---
title: View.ShowParagraphs Property (Word)
keywords: vbawd10.chm161808401
f1_keywords:
- vbawd10.chm161808401
ms.prod: word
api_name:
- Word.View.ShowParagraphs
ms.assetid: 17b2ea55-14d3-1606-1d45-da601009a209
ms.date: 06/08/2017
---


# View.ShowParagraphs Property (Word)

 **True** if paragraph marks are displayed. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowParagraphs**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Example

This example hides paragraph marks in the active window.


```vb
ActiveDocument.ActiveWindow.View.ShowParagraphs = False
```


## See also


#### Concepts


[View Object](view-object-word.md)

