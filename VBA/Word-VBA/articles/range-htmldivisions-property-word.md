---
title: Range.HTMLDivisions Property (Word)
keywords: vbawd10.chm157155734
f1_keywords:
- vbawd10.chm157155734
ms.prod: word
api_name:
- Word.Range.HTMLDivisions
ms.assetid: 4712d81f-7028-357b-a7ff-dc4f382cc5e3
ms.date: 06/08/2017
---


# Range.HTMLDivisions Property (Word)

Returns an  **HTMLDivisions** object that represents an HTML division in a Web document.


## Syntax

 _expression_ . **HTMLDivisions**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Example

This example formats three nested divisions in the active document. This example assumes that the active document is an HTML document with at least three divisions.


```vb
Sub FormatHTMLDivisions() 
 With ActiveDocument.Range.HTMLDivisions(1) 
 With .Borders(wdBorderLeft) 
 .Color = wdColorRed 
 .LineStyle = wdLineStyleSingle 
 End With 
 With .Borders(wdBorderRight) 
 .Color = wdColorRed 
 .LineStyle = wdLineStyleSingle 
 End With 
 With .HTMLDivisions(1) 
 .LeftIndent = InchesToPoints(1) 
 .RightIndent = InchesToPoints(1) 
 With .Borders(wdBorderTop) 
 .Color = wdColorBlue 
 .LineStyle = wdLineStyleDouble 
 End With 
 With .Borders(wdBorderBottom) 
 .Color = wdColorBlue 
 .LineStyle = wdLineStyleDouble 
 End With 
 With .HTMLDivisions(1) 
 .LeftIndent = InchesToPoints(1) 
 .RightIndent = InchesToPoints(1) 
 With .Borders(wdBorderLeft) 
 .LineStyle = wdLineStyleDot 
 End With 
 With .Borders(wdBorderRight) 
 .LineStyle = wdLineStyleDot 
 End With 
 With .Borders(wdBorderTop) 
 .LineStyle = wdLineStyleDot 
 End With 
 With .Borders(wdBorderBottom) 
 .LineStyle = wdLineStyleDot 
 End With 
 End With 
 End With 
 End With 
 
End Sub
```


## See also


#### Concepts


[Range Object](range-object-word.md)

