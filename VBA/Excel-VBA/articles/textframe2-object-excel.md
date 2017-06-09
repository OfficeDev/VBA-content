---
title: TextFrame2 Object (Excel)
ms.prod: excel
api_name:
- Excel.TextFrame2
ms.assetid: 66ba23e5-9b15-b954-a1db-1bd19b4eb90d
ms.date: 06/08/2017
---


# TextFrame2 Object (Excel)

Represents the text frame in a  **[Shape](shape-object-excel.md)**, **[ShapeRange](shaperange-object-excel.md)**, or **[ChartFormat](chartformat-object-excel.md)** object.


## Remarks

This object contains the text in the text frame as well as the properties and methods that control the alignment and anchoring of the text frame. Use the  **TextFrame2** property to return a **TextFrame2** object.


## Example

The following example adds a rectangle to  `myDocument`, adds text to the rectangle, and then sets the margins for the text frame.


```
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 0, 0, 250, 140).TextFrame2 
 .TextRange.Text = "Here is some test text" 
 .MarginBottom = 10 
 .MarginLeft = 10 
 .MarginRight = 10 
 .MarginTop = 10 
End With
```


## Methods



|**Name**|
|:-----|
|[DeleteText](textframe2-deletetext-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](textframe2-application-property-excel.md)|
|[AutoSize](textframe2-autosize-property-excel.md)|
|[Column](textframe2-column-property-excel.md)|
|[Creator](textframe2-creator-property-excel.md)|
|[HasText](textframe2-hastext-property-excel.md)|
|[HorizontalAnchor](textframe2-horizontalanchor-property-excel.md)|
|[MarginBottom](textframe2-marginbottom-property-excel.md)|
|[MarginLeft](textframe2-marginleft-property-excel.md)|
|[MarginRight](textframe2-marginright-property-excel.md)|
|[MarginTop](textframe2-margintop-property-excel.md)|
|[NoTextRotation](textframe2-notextrotation-property-excel.md)|
|[Orientation](textframe2-orientation-property-excel.md)|
|[Parent](textframe2-parent-property-excel.md)|
|[PathFormat](textframe2-pathformat-property-excel.md)|
|[Ruler](textframe2-ruler-property-excel.md)|
|[TextRange](textframe2-textrange-property-excel.md)|
|[ThreeD](textframe2-threed-property-excel.md)|
|[VerticalAnchor](textframe2-verticalanchor-property-excel.md)|
|[WarpFormat](textframe2-warpformat-property-excel.md)|
|[WordArtformat](textframe2-wordartformat-property-excel.md)|
|[WordWrap](textframe2-wordwrap-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
