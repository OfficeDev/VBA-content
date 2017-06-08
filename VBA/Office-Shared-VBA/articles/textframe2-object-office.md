---
title: TextFrame2 Object (Office)
ms.prod: office
api_name:
- Office.TextFrame2
ms.assetid: d2903007-70d4-0b98-e617-96fb2df26975
ms.date: 06/08/2017
---


# TextFrame2 Object (Office)

Represents the text frame in a  **Shape** or **ShapeRange** object. Contains the text in the text frame and exposes properties and methods that control the alignment and anchoring of the text frame.


## Remarks

Use the TextFrame2 property of the Shape and ShapeRange objects to return a TextFrame2 object. 


## Example

The following code adds a rectangle to a slide, adds text to the rectangle, and then sets the margins for the text frame.


```
Set pptSlide = ActivePresentation.Slides(1) 
With pptSlide.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame2 
 .TextRange.Text = "Here is some sample text" 
 .MarginBottom = 10 
 .MarginLeft = 10 
 .MarginRight = 10 
 .MarginTop = 10 
End With 

```


## Methods



|**Name**|
|:-----|
|[DeleteText](textframe2-deletetext-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](textframe2-application-property-office.md)|
|[AutoSize](textframe2-autosize-property-office.md)|
|[Column](textframe2-column-property-office.md)|
|[Creator](textframe2-creator-property-office.md)|
|[HasText](textframe2-hastext-property-office.md)|
|[HorizontalAnchor](textframe2-horizontalanchor-property-office.md)|
|[MarginBottom](textframe2-marginbottom-property-office.md)|
|[MarginLeft](textframe2-marginleft-property-office.md)|
|[MarginRight](textframe2-marginright-property-office.md)|
|[MarginTop](textframe2-margintop-property-office.md)|
|[NoTextRotation](textframe2-notextrotation-property-office.md)|
|[Orientation](textframe2-orientation-property-office.md)|
|[Parent](textframe2-parent-property-office.md)|
|[PathFormat](textframe2-pathformat-property-office.md)|
|[Ruler](textframe2-ruler-property-office.md)|
|[TextRange](textframe2-textrange-property-office.md)|
|[ThreeD](textframe2-threed-property-office.md)|
|[VerticalAnchor](textframe2-verticalanchor-property-office.md)|
|[WarpFormat](textframe2-warpformat-property-office.md)|
|[WordArtformat](textframe2-wordartformat-property-office.md)|
|[WordWrap](textframe2-wordwrap-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
