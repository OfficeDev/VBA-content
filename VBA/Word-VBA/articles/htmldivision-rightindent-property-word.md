---
title: HTMLDivision.RightIndent Property (Word)
keywords: vbawd10.chm166133764
f1_keywords:
- vbawd10.chm166133764
ms.prod: word
api_name:
- Word.HTMLDivision.RightIndent
ms.assetid: d691b48c-343f-5b4a-666b-83cae994b8b9
ms.date: 06/08/2017
---


# HTMLDivision.RightIndent Property (Word)

Returns or sets the right indent (in points) for the specified paragraphs. Read/write  **Single** .


## Syntax

 _expression_ . **RightIndent**

 _expression_ Required. A variable that represents an **[HTMLDivision](htmldivision-object-word.md)** object.


## Example

This example sets the right indent for all paragraphs in the active document to 1 inch from the right margin. The  **InchesToPoints** method is used to convert inches to points.


```vb
ActiveDocument.Paragraphs.RightIndent = InchesToPoints(1)
```


## See also


#### Concepts


[HTMLDivision Object](htmldivision-object-word.md)

