---
title: WrapFormat Object (Publisher)
keywords: vbapb10.chm851967
f1_keywords:
- vbapb10.chm851967
ms.prod: publisher
api_name:
- Publisher.WrapFormat
ms.assetid: b6f80d40-2043-6944-3ed8-f26635c7fa4d
ms.date: 06/08/2017
---


# WrapFormat Object (Publisher)

Represents all the properties for wrapping text around a shape or shape range.
 


## Example

Use the  **[TextWrap](shape-textwrap-property-publisher.md)** property to return a **WrapFormat** object. The following example adds an oval to the active publication and specifies that publication text wrap around the left and right sides of the square that circumscribes the oval. There will be a 0.1-inch margin between the publication text and the top, bottom, left side, and right side of the square.
 

 

```
Sub SetTextWrapFormatProperties() 
 Dim shpOval As Shape 
 
 Set shpOval = ActiveDocument.Pages(1).Shapes.AddShape(Type:=msoShapeOval, _ 
 Left:=36, Top:=36, Width:=100, Height:=35) 
 With shpOval.TextWrap 
 .Type = pbWrapTypeSquare 
 .Side = pbWrapSideBoth 
 .DistanceAuto = msoFalse 
 .DistanceTop = InchesToPoints(0.1) 
 .DistanceBottom = InchesToPoints(0.1) 
 .DistanceLeft = InchesToPoints(0.1) 
 .DistanceRight = InchesToPoints(0.1) 
 End With 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](wrapformat-application-property-publisher.md)|
|[DistanceAuto](wrapformat-distanceauto-property-publisher.md)|
|[DistanceBottom](wrapformat-distancebottom-property-publisher.md)|
|[DistanceLeft](wrapformat-distanceleft-property-publisher.md)|
|[DistanceRight](wrapformat-distanceright-property-publisher.md)|
|[DistanceTop](wrapformat-distancetop-property-publisher.md)|
|[Parent](wrapformat-parent-property-publisher.md)|
|[Side](wrapformat-side-property-publisher.md)|
|[Type](wrapformat-type-property-publisher.md)|

