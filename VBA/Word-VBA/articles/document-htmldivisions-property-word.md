---
title: Document.HTMLDivisions Property (Word)
keywords: vbawd10.chm158007638
f1_keywords:
- vbawd10.chm158007638
ms.prod: word
api_name:
- Word.Document.HTMLDivisions
ms.assetid: 8e383427-0777-116c-12d8-59bcc3f819d1
ms.date: 06/08/2017
---


# Document.HTMLDivisions Property (Word)

Returns an  **[HTMLDivisions](htmldivisions-object-word.md)** collection that represents the HTML DIV elements in a Web document.


## Syntax

 _expression_ . **HTMLDivisions**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example formats three nested divisions in the active document. This example assumes that the active document is an HTML document with at least three divisions.


```vb
Sub FormatHTMLDivisions() 
 With ActiveDocument.HTMLDivisions(1) 
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


[Document Object](document-object-word.md)

