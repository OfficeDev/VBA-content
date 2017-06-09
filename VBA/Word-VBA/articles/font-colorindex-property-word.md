---
title: Font.ColorIndex Property (Word)
keywords: vbawd10.chm156369033
f1_keywords:
- vbawd10.chm156369033
ms.prod: word
api_name:
- Word.Font.ColorIndex
ms.assetid: c5011017-bf7a-5d89-0f20-f000d3ffd0ea
ms.date: 06/08/2017
---


# Font.ColorIndex Property (Word)

Returns or sets a  **WdColorIndex** constant that represents the color for the specified font. Read/write .


## Syntax

 _expression_ . **ColorIndex**

 _expression_ Required. A variable that represents a **[Font](font-object-word.md)** object.


## Remarks

The  **wdByAuthor** constant is not a valid color for fonts.


## Example

This example changes the color of the text in the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).Range.Font.ColorIndex = wdGreen
```

This example formats the selected text to appear in red.




```
Selection.Font.ColorIndex = wdRed
```


## See also


#### Concepts


[Font Object](font-object-word.md)

