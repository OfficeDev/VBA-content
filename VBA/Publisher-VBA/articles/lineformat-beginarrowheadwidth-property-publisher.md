---
title: LineFormat.BeginArrowheadWidth Property (Publisher)
keywords: vbapb10.chm3408131
f1_keywords:
- vbapb10.chm3408131
ms.prod: publisher
api_name:
- Publisher.LineFormat.BeginArrowheadWidth
ms.assetid: a752c674-1b83-b8c8-d325-b61804f5fadc
ms.date: 06/08/2017
---


# LineFormat.BeginArrowheadWidth Property (Publisher)

Returns or sets an  **MsoArrowheadWidth**constant indicating the width of the arrowhead at the beginning of the specified line. Read/write.


## Syntax

 _expression_. **BeginArrowheadWidth**

 _expression_A variable that represents a  **LineFormat** object.


### Return Value

MsoArrowheadWidth


## Remarks

The  **BeginArrowheadWidth** property value can be one of the ** [MsoArrowheadWidth](http://msdn.microsoft.com/library/7183f2e0-7431-170b-f4e7-3f8737017ed8%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.

Use the  **[EndArrowheadWidth](lineformat-endarrowheadwidth-property-publisher.md)** property to return or set the width of the arrowhead at the end of the line.


## Example

This example adds a line to the active publication. There is a short, narrow oval on the line's starting point and a long, wide triangle on its endpoint.


```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddLine(BeginX:=100, BeginY:=100, _ 
 EndX:=200, EndY:=300).Line 
 .BeginArrowheadLength = msoArrowheadShort 
 .BeginArrowheadStyle = msoArrowheadOval 
 .BeginArrowheadWidth = msoArrowheadNarrow 
 .EndArrowheadLength = msoArrowheadLong 
 .EndArrowheadStyle = msoArrowheadTriangle 
 .EndArrowheadWidth = msoArrowheadWide 
End With 

```


