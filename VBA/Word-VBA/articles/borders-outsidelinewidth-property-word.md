---
title: Borders.OutsideLineWidth Property (Word)
keywords: vbawd10.chm154927113
f1_keywords:
- vbawd10.chm154927113
ms.prod: word
api_name:
- Word.Borders.OutsideLineWidth
ms.assetid: 4f2c2f67-7a0e-e06a-c49c-30e8446bebe2
ms.date: 06/08/2017
---


# Borders.OutsideLineWidth Property (Word)

Returns or sets the line width of the outside border of an object. Read/write.


## Syntax

 _expression_ . **OutsideLineWidth**

 _expression_ Required. A variable that represents a **[Borders](borders-object-word.md)** collection.


## Remarks

This property returns  **wdUndefined** if the object has outside borders with more than one line width; otherwise, returns **False** or a **WdLineWidth** constant. Can be set to **True** , **False** , or a **WdLineWidth** constant.


## Example

This example adds a wavy border around the first table in the active document.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 With ActiveDocument.Tables(1).Borders 
 .OutsideLineStyle = wdLineStyleSingleWavy 
 .OutsideLineWidth = wdLineWidth075pt 
 End With 
End If
```

This example adds dotted borders around the first four paragraphs in the active document.




```vb
Set myDoc = ActiveDocument 
Set myRange = myDoc.Range(Start:=myDoc.Paragraphs(1).Range.Start, _ 
 End:=myDoc.Paragraphs(4).Range.End) 
myRange.Borders.OutsideLineStyle = wdLineStyleDot 
myRange.Borders.OutsideLineWidth = wdLineWidth075pt
```


## See also


#### Concepts


[Borders Collection Object](borders-object-word.md)

