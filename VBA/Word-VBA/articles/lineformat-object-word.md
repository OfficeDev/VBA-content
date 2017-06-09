---
title: LineFormat Object (Word)
keywords: vbawd10.chm2506
f1_keywords:
- vbawd10.chm2506
ms.prod: word
api_name:
- Word.LineFormat
ms.assetid: 28fabccb-d03f-3466-9d07-ea3ebc4cdd11
ms.date: 06/08/2017
---


# LineFormat Object (Word)

Represents line and arrowhead formatting. For a line, the  **LineFormat** object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border.


## Remarks

Use the  **Line** property to return a **LineFormat** object. The following example adds a blue, dashed line to the active document. There is a short, narrow oval at the line's starting point and a long, wide triangle at its endpoint.


```vb
With ActiveDocument.Shapes.AddLine(100, 100, 200, 300).Line 
 .DashStyle = msoLineDashDotDot 
 .ForeColor.RGB = RGB(50, 0, 128) 
 .BeginArrowheadLength = msoArrowheadShort 
 .BeginArrowheadStyle = msoArrowheadOval 
 .BeginArrowheadWidth = msoArrowheadNarrow 
 .EndArrowheadLength = msoArrowheadLong 
 .EndArrowheadStyle = msoArrowheadTriangle 
 .EndArrowheadWidth = msoArrowheadWide 
End With
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

