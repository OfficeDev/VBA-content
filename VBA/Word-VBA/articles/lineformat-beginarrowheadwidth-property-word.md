---
title: LineFormat.BeginArrowheadWidth Property (Word)
keywords: vbawd10.chm164233319
f1_keywords:
- vbawd10.chm164233319
ms.prod: word
api_name:
- Word.LineFormat.BeginArrowheadWidth
ms.assetid: f15fdfd3-dd6c-a47e-8fad-ee8367c72341
ms.date: 06/08/2017
---


# LineFormat.BeginArrowheadWidth Property (Word)

Returns or sets the width of the arrowhead at the beginning of the specified line. Read/write  **MsoArrowheadWidth** .


## Syntax

 _expression_ . **BeginArrowheadWidth**

 _expression_ Required. A variable that represents a **[LineFormat](lineformat-object-word.md)** object.


## Example

This example adds a line to the first document. There is a short, narrow oval on the line's starting point and a long, wide triangle on its endpoint.


```vb
Dim docFirst As Document 
 
Set docFirst =
```


```vb
Documents(1) 
With docFirst.Shapes.AddLine(100, 100, 200, 300).Line 
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

