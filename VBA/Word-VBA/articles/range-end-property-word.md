---
title: Range.End Property (Word)
keywords: vbawd10.chm157155332
f1_keywords:
- vbawd10.chm157155332
ms.prod: word
api_name:
- Word.Range.End
ms.assetid: fe90f321-c7b5-bea2-fa60-e6b750b33cf7
ms.date: 06/08/2017
---


# Range.End Property (Word)

Returns or sets the ending character position of a range. Read/write  **Long** .


## Syntax

 _expression_ . **End**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

 **Range** objects all have a starting position and an ending position. The ending position is the point farthest away from the beginning of the story. If this property is set to a value smaller than the **[Start](range-start-property-word.md)** property, the **Start** property is set to the same value (that is, the **Start** and **End** property are equal).

This property returns the ending character position relative to the beginning of the story. The main document story ( **wdMainTextStory** ) begins with character position 0 (zero). You can change the size of a selection, range, or bookmark by setting this property.


## Example

This example changes the ending position of myRange by one character.


```vb
Set myRange = ActiveDocument.Paragraphs(1).Range 
myRange.End = myRange.End - 1
```


## See also


#### Concepts


[Range Object](range-object-word.md)

