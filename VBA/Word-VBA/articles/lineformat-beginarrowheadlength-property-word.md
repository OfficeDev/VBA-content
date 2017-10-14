---
title: LineFormat.BeginArrowheadLength Property (Word)
keywords: vbawd10.chm164233317
f1_keywords:
- vbawd10.chm164233317
ms.prod: word
api_name:
- Word.LineFormat.BeginArrowheadLength
ms.assetid: e2bcb274-001e-69a8-35de-009193dcc117
ms.date: 06/08/2017
---


# LineFormat.BeginArrowheadLength Property (Word)

Returns or sets the length of the arrowhead at the beginning of the specified line. Read/write  **MsoArrowheadLength** .


## Syntax

 _expression_ . **BeginArrowheadLength**

 _expression_ Required. A variable that represents a **[LineFormat](lineformat-object-word.md)** object.


## Example

This example adds a line to the active document. There is a short, narrow oval on the line's starting point and a long, wide triangle on its endpoint.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 
 
With docActive.Shapes.AddLine(100, 100, 200, 300).Line 
 .BeginArrowheadLength = msoArrowheadShort 
 .BeginArrowheadStyle = msoArrowheadOval 
 .BeginArrowheadWidth = msoArrowheadNarrow 
 .EndArrowheadLength = msoArrowheadLong 
 .EndArrowheadStyle = msoArrowheadTriangle 
 .EndArrowheadWidth = msoArrowheadWide 
End With
```


## See also


#### Concepts


[LineFormat Object](lineformat-object-word.md)

