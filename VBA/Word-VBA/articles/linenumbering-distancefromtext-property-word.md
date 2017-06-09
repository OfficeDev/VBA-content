---
title: LineNumbering.DistanceFromText Property (Word)
keywords: vbawd10.chm158466150
f1_keywords:
- vbawd10.chm158466150
ms.prod: word
api_name:
- Word.LineNumbering.DistanceFromText
ms.assetid: cc541a06-5216-1a7a-9db1-172c94272d31
ms.date: 06/08/2017
---


# LineNumbering.DistanceFromText Property (Word)

Returns or sets the distance (in points) between the right edge of line numbers and the left edge of the document text. Read/write  **Single** .


## Syntax

 _expression_ . **DistanceFromText**

 _expression_ A variable that represents a **[LineNumbering](linenumbering-object-word.md)** object.


## Example

This example adds line numbers to the active document. The distance between the line numbers and the left margin is 36 points (0.5 inch).


```vb
With ActiveDocument.PageSetup.LineNumbering 
 .Active = True 
 .CountBy = 5 
 .DistanceFromText = 36 
End With
```


## See also


#### Concepts


[LineNumbering Object](linenumbering-object-word.md)

