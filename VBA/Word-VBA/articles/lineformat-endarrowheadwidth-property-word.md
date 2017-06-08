---
title: LineFormat.EndArrowheadWidth Property (Word)
keywords: vbawd10.chm164233323
f1_keywords:
- vbawd10.chm164233323
ms.prod: word
api_name:
- Word.LineFormat.EndArrowheadWidth
ms.assetid: 01d77438-aa35-983b-7d93-a88e135d1820
ms.date: 06/08/2017
---


# LineFormat.EndArrowheadWidth Property (Word)

Returns or sets the width of the arrowhead at the end of the specified line. Read/write  **MsoArrowheadWidth** .


## Syntax

 _expression_ . **EndArrowheadWidth**

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

