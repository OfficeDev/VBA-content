---
title: ParagraphFormat.Tabs Property (Publisher)
keywords: vbapb10.chm5439506
f1_keywords:
- vbapb10.chm5439506
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.Tabs
ms.assetid: c42ba898-b84f-7215-129d-8134670f75ac
ms.date: 06/08/2017
---


# ParagraphFormat.Tabs Property (Publisher)

Returns a  **[TabStops](tabstops-object-publisher.md)** object representing the custom and default tabs for a paragraph or group of paragraphs.


## Syntax

 _expression_. **Tabs**

 _expression_A variable that represents a  **ParagraphFormat** object.


### Return Value

TabStops


## Example

The following example adds two tab stops to the selected paragraphs. The first tab stop is a left-aligned tab with a dotted tab leader positioned at 1 inch (72 points). The second tab stop is centered and is positioned at 2 inches.


```vb
Dim tabsAll As TabStops 
 
Set tabsAll = Selection.TextRange.ParagraphFormat.Tabs 
 
With tabsAll 
 .Add Position:=InchesToPoints(1), _ 
 Leader:=pbTabLeaderDot, Alignment:=pbTabAlignmentLeading 
 .Add Position:=InchesToPoints(2), _ 
 Leader:=pbTabLeaderNone, Alignment:=pbTabAlignmentCenter 
End With
```


