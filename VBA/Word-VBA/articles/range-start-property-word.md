---
title: Range.Start Property (Word)
keywords: vbawd10.chm157155331
f1_keywords:
- vbawd10.chm157155331
ms.prod: word
api_name:
- Word.Range.Start
ms.assetid: aadedbb7-1ee2-9e5a-296d-0ebe25b6d8f4
ms.date: 06/08/2017
---


# Range.Start Property (Word)

Returns or sets the starting character position of a range. Read/write  **Long** .


## Syntax

 _expression_ . **Start**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

 **Range** objects have starting and ending character positions. The starting position refers to the character position closest to the beginning of the story. If this property is set to a value larger than that of the **End** property, the **End** property is set to the same value as that of **Start** property.

This property returns the starting character position relative to the beginning of the story. The main text story ( **wdMainTextStory** ) begins with character position 0 (zero). You can change the size of a selection, range, or bookmark by setting this property.


## Example

This example returns the starting position of the second paragraph and the ending position of the fourth paragraph in the active document. The character positions are used to create the range myRange.


```vb
pos = ActiveDocument.Paragraphs(2).Range.Start 
pos2 = ActiveDocument.Paragraphs(4).Range.End 
Set myRange = ActiveDocument.Range(Start:=pos, End:=pos2)
```

This example moves the starting position of myRange one character to the right (this reduces the size of the range by one character).




```vb
Set myRange = Selection.Range 
myRange.SetRange Start:=myRange.Start + 1, End:=myRange.End
```


## See also


#### Concepts


[Range Object](range-object-word.md)

