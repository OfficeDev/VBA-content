---
title: ListFormat.ListOutdent Method (Word)
keywords: vbawd10.chm163578066
f1_keywords:
- vbawd10.chm163578066
ms.prod: word
api_name:
- Word.ListFormat.ListOutdent
ms.assetid: f69834f5-ae8b-f67a-a5b5-131a7382b1c5
ms.date: 06/08/2017
---


# ListFormat.ListOutdent Method (Word)

Decreases the list level of the paragraphs in the range for the specified  **ListFormat** object, in increments of one level.


## Syntax

 _expression_ . **ListOutdent**

 _expression_ Required. A variable that represents a **[ListFormat](listformat-object-word.md)** object.


## Example

This example reduces the indent of each paragraph in first list in the active document by one level.


```vb
ActiveDocument.Lists(1).Range.ListFormat.ListOutdent
```

This example formats paragraphs four through eight in the active document as an outline-numbered list, indents the paragraphs one level, and then removes the indent from the first paragraph in the list.




```vb
Dim docActive As Document 
Dim rngTemp As Range 
 
Set docActive = ActiveDocument
```




```vb
Set rngTemp = _ 
 docActive.Range( _ 
 Start:=docActive.Paragraphs(4).Range.Start, _ 
 End:=docActive.Paragraphs(8).Range.End) 
 
With rngTemp.ListFormat 
 .ApplyOutlineNumberDefault 
 .ListIndent 
End With 
 
docActive.Paragraphs(4).Range.ListFormat.ListOutdent
```


## See also


#### Concepts


[ListFormat Object](listformat-object-word.md)

