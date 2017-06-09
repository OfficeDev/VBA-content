---
title: LineFormat Object (PowerPoint)
keywords: vbapp10.chm553000
f1_keywords:
- vbapp10.chm553000
ms.prod: powerpoint
api_name:
- PowerPoint.LineFormat
ms.assetid: 11c955d5-bbda-d99f-cec9-fc6187450a12
ms.date: 06/08/2017
---


# LineFormat Object (PowerPoint)

Represents line and arrowhead formatting. For a line, the  **LineFormat** object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border.


## Example

Use the  **Line** property to return a **LineFormat** object. The following example adds a blue, dashed line to `myDocument`. There's a short, narrow oval at the line's starting point and a long, wide triangle at its endpoint.


```
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddLine(100, 100, 200, 300).Line

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


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/e1f2c525-54c7-80f3-5f80-bca0a9e0a63c%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/5c8e915a-6fb6-92b1-1d49-a74ee3a3e06d%28Office.15%29.aspx)|
|[BeginArrowheadLength](http://msdn.microsoft.com/library/b46151e1-251f-7498-9dfc-b652b356edf0%28Office.15%29.aspx)|
|[BeginArrowheadStyle](http://msdn.microsoft.com/library/04f6e7f1-c76f-b70d-5fbd-daaa907fe59d%28Office.15%29.aspx)|
|[BeginArrowheadWidth](http://msdn.microsoft.com/library/3834e2c8-d153-57f8-014e-1545326dd370%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/e4020bf2-0b36-4e77-3850-949ac81e0c86%28Office.15%29.aspx)|
|[DashStyle](http://msdn.microsoft.com/library/7fc898b4-1eea-21fc-52e5-0ec92bde527f%28Office.15%29.aspx)|
|[EndArrowheadLength](http://msdn.microsoft.com/library/e7e183f6-fc85-0a5f-c1c1-f182c8020c20%28Office.15%29.aspx)|
|[EndArrowheadStyle](http://msdn.microsoft.com/library/8f4f7a0a-cbfa-ee6c-25bb-b1aca1e2b883%28Office.15%29.aspx)|
|[EndArrowheadWidth](http://msdn.microsoft.com/library/5830e4ff-c630-198a-ea2b-b5d1397ea846%28Office.15%29.aspx)|
|[ForeColor](http://msdn.microsoft.com/library/0b022f2e-d546-2d56-13ae-1040682ee9d0%28Office.15%29.aspx)|
|[InsetPen](http://msdn.microsoft.com/library/07a69459-0a24-c9b8-5aba-103b39d8b1af%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/6644560e-0d3c-d675-b8a0-3481496c12ec%28Office.15%29.aspx)|
|[Pattern](http://msdn.microsoft.com/library/5c4c7e5a-1932-01a4-034d-0a4e98c43174%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/8a9b1a85-f290-97f5-c19d-6427d1214f7b%28Office.15%29.aspx)|
|[Transparency](http://msdn.microsoft.com/library/7d9e3a3c-479a-1a7a-45b2-4245b8444c21%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/4b10ecb4-01f1-019f-62f8-2a3508a01ca3%28Office.15%29.aspx)|
|[Weight](http://msdn.microsoft.com/library/5141d66f-4706-060d-fb4c-f244f9ac6437%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
