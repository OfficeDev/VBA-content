---
title: ShadowFormat Object (Publisher)
keywords: vbapb10.chm3735551
f1_keywords:
- vbapb10.chm3735551
ms.prod: publisher
api_name:
- Publisher.ShadowFormat
ms.assetid: b23ab92e-5e49-8d8d-69d5-93d391a9edb2
ms.date: 06/08/2017
---


# ShadowFormat Object (Publisher)

Represents shadow formatting for a shape.
 


## Example

Use the  **Shadow** property to return a **ShadowFormat** object. The following example adds a shadowed rectangle to the active document. The pink shadow is offset 7 points to the right of the rectangle and 7 points above it.
 

 

```
Sub FormatShadow() 
 With ActiveDocument.Pages(1).Shapes.AddShape( _ 
 Type:=msoShapeRectangle, Left:=72, Top:=72, _ 
 Width:=100, Height:=200).Shadow 
 .ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=150) 
 .Obscured = msoTrue 
 .OffsetX = 7 
 .OffsetY = -7 
 .Visible = True 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[IncrementOffsetX](shadowformat-incrementoffsetx-method-publisher.md)|
|[IncrementOffsetY](shadowformat-incrementoffsety-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](shadowformat-application-property-publisher.md)|
|[Blur](shadowformat-blur-property-publisher.md)|
|[ForeColor](shadowformat-forecolor-property-publisher.md)|
|[Obscured](shadowformat-obscured-property-publisher.md)|
|[OffsetX](shadowformat-offsetx-property-publisher.md)|
|[OffsetY](shadowformat-offsety-property-publisher.md)|
|[Parent](shadowformat-parent-property-publisher.md)|
|[RotateWithShape](shadowformat-rotatewithshape-property-publisher.md)|
|[Size](shadowformat-size-property-publisher.md)|
|[Type](shadowformat-type-property-publisher.md)|
|[Visible](shadowformat-visible-property-publisher.md)|

