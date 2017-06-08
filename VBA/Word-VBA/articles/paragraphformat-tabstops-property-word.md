---
title: ParagraphFormat.TabStops Property (Word)
keywords: vbawd10.chm156435535
f1_keywords:
- vbawd10.chm156435535
ms.prod: word
api_name:
- Word.ParagraphFormat.TabStops
ms.assetid: 9eed85b9-aee6-04af-c5ce-f6ba47176d35
ms.date: 06/08/2017
---


# ParagraphFormat.TabStops Property (Word)

Returns or sets a  **TabStops** collection that represents all the custom tab stops for the specified paragraphs. Read/write.


## Syntax

 _expression_ . **TabStops**

 _expression_ A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example adds a centered tab stop at 2 inches to all the paragraphs in the active document. The  **InchesToPoints** method is used to convert inches to points.


```vb
With ActiveDocument.Paragraphs.TabStops 
 .Add Position:= InchesToPoints(2), Alignment:= wdAlignTabCenter 
End With
```

This example sets the tab stops for every paragraph in the document to match the tab stops in the first paragraph.




```vb
Set para1Tabs = ActiveDocument.Paragraphs(1).TabStops 
ActiveDocument.Paragraphs.TabStops = para1Tabs
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)

