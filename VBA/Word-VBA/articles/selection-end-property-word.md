---
title: Selection.End Property (Word)
keywords: vbawd10.chm158662660
f1_keywords:
- vbawd10.chm158662660
ms.prod: word
api_name:
- Word.Selection.End
ms.assetid: 99e3bd79-a8f1-8057-1ac2-b9e76eab99ff
ms.date: 06/08/2017
---


# Selection.End Property (Word)

Returns or sets the ending character position of a selection. Read/write  **Long** .


## Syntax

 _expression_ . **End**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

If this property is set to a value smaller than the  **Start** property, the **Start** property is set to the same value (that is, the **Start** and **End** property are equal).

The  **Selection** object has a starting position and an ending position. The ending position is the point farthest away from the beginning of the story. This property returns the ending character position relative to the beginning of the story. The main document story ( **wdMainTextStory** ) begins with character position 0 (zero). You can change the size of a selection by setting this property.


## Example

This example retrieves the ending position of the selection. This value is used to create a range so that a field can be inserted after the selection.


```vb
pos = Selection.End 
Set myRange = ActiveDocument.Range(Start:=pos, End:=pos) 
ActiveDocument.Fields.Add Range:=myRange, Type:=wdFieldAuthor
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

