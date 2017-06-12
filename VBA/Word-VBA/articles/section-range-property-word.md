---
title: Section.Range Property (Word)
keywords: vbawd10.chm156827648
f1_keywords:
- vbawd10.chm156827648
ms.prod: word
api_name:
- Word.Section.Range
ms.assetid: aabe43c6-4ace-af98-f773-bc547f289c9d
ms.date: 06/08/2017
---


# Section.Range Property (Word)

Returns a  **Range** object that represents the portion of a document that's contained in the specified object.


## Syntax

 _expression_ . **Range**

 _expression_ Required. A variable that represents a **[Section](section-object-word.md)** object.


## Example

This example inserts text at the end of section one.


```vb
Set myRange = ActiveDocument.Sections(1).Range 
With myRange 
 .MoveEnd Unit:=wdCharacter, Count:=-1 
 .Collapse Direction:=wdCollapseEnd 
 .InsertParagraphAfter 
 .InsertAfter "End of section" 
End With
```


## See also


#### Concepts


[Section Object](section-object-word.md)

