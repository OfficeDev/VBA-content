---
title: Borders.SurroundHeader Property (Word)
keywords: vbawd10.chm154927128
f1_keywords:
- vbawd10.chm154927128
ms.prod: word
api_name:
- Word.Borders.SurroundHeader
ms.assetid: bada52f5-7f73-8565-bd7b-33311a1aa942
ms.date: 06/08/2017
---


# Borders.SurroundHeader Property (Word)

 **True** if a page border encompasses the document header. Read/write **Boolean** .


## Syntax

 _expression_ . **SurroundHeader**

 _expression_ An expression that returns a **[Borders](borders-object-word.md)** collection object.


## Example

This example formats the page border in section one of the active document to exclude the header and footer areas on each page.


```vb
With ActiveDocument.Sections(1).Borders 
 .SurroundFooter = False 
 .SurroundHeader = False 
End With
```


## See also


#### Concepts


[Borders Collection Object](borders-object-word.md)

