---
title: Borders.InsideColor Property (Word)
keywords: vbawd10.chm154927136
f1_keywords:
- vbawd10.chm154927136
ms.prod: word
api_name:
- Word.Borders.InsideColor
ms.assetid: 90205db3-2b44-25dd-3b61-b9dc3ccb157a
ms.date: 06/08/2017
---


# Borders.InsideColor Property (Word)

Returns or sets the 24-bit color of the inside borders. Read/write.


## Syntax

 _expression_ . **InsideColor**

 _expression_ Required. A variable that represents a **[Borders](borders-object-word.md)** collection.


## Remarks

This property can be any valid  **WdColor** constant or a value returned by Visual Basic's **RGB** function. If the **InsideLineStyle** property is set to either **wdLineStyleNone** or **False** , setting this property has no effect.


## Example

This example adds dark red borders between the first four paragraphs in the active document.


```vb
Set myDoc = ActiveDocument 
Set myRange = myDoc.Range(Start:=myDoc.Paragraphs(1).Range.Start, _ 
 End:=myDoc.Paragraphs(4).Range.End) 
With myRange.Borders 
 .InsideLineStyle = wdLineStyleSingle 
 .InsideLineWidth = wdLineWidth150pt 
 .InsideColor = wdDarkRed 
End With
```


## See also


#### Concepts


[Borders Collection Object](borders-object-word.md)

