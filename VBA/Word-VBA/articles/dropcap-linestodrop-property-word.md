---
title: DropCap.LinesToDrop Property (Word)
keywords: vbawd10.chm156631052
f1_keywords:
- vbawd10.chm156631052
ms.prod: word
api_name:
- Word.DropCap.LinesToDrop
ms.assetid: 148ca844-d9ee-39f5-722a-6bd8279ca4b9
ms.date: 06/08/2017
---


# DropCap.LinesToDrop Property (Word)

Returns or sets the height (in lines) of the specified dropped capital letter. Read/write  **Long** .


## Syntax

 _expression_ . **LinesToDrop**

 _expression_ An expression that returns a **[DropCap](dropcap-object-word.md)** object.


## Example

This example formats the first character in the active document as a dropped capital letter with a height of three lines.


```vb
With ActiveDocument.Paragraphs(1).DropCap 
 .Enable 
 .Position = wdDropNormal 
 .LinesToDrop = 3 
End With
```


## See also


#### Concepts


[DropCap Object](dropcap-object-word.md)

