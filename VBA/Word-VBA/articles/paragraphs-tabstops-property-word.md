---
title: Paragraphs.TabStops Property (Word)
keywords: vbawd10.chm156763215
f1_keywords:
- vbawd10.chm156763215
ms.prod: word
api_name:
- Word.Paragraphs.TabStops
ms.assetid: cf369030-7569-699f-d8be-7a24b63e22eb
ms.date: 06/08/2017
---


# Paragraphs.TabStops Property (Word)

Returns or sets a  **TabStops** collection that represents all the custom tab stops for the specified paragraphs. Read/write.


## Syntax

 _expression_ . **TabStops**

 _expression_ A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


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


[Paragraphs Collection Object](paragraphs-object-word.md)

