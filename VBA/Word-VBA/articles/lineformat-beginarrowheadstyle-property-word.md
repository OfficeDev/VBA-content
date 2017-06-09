---
title: LineFormat.BeginArrowheadStyle Property (Word)
keywords: vbawd10.chm164233318
f1_keywords:
- vbawd10.chm164233318
ms.prod: word
api_name:
- Word.LineFormat.BeginArrowheadStyle
ms.assetid: 16aa1b91-5126-bbe5-be7d-ce26245f50a2
ms.date: 06/08/2017
---


# LineFormat.BeginArrowheadStyle Property (Word)

Returns or sets the style of the arrowhead at the beginning of the specified line. Read/write  **MsoArrowheadStyle** .


## Syntax

 _expression_ . **BeginArrowheadStyle**

 _expression_ Required. A variable that represents a **[LineFormat](lineformat-object-word.md)** object.


## Example

This example adds a line to the active document. There is a short, narrow oval on the line's starting point and a long, wide triangle on its endpoint.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 

```


```vb
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

