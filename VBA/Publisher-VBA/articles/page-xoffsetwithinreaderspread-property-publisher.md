---
title: Page.XOffsetWithinReaderSpread Property (Publisher)
keywords: vbapb10.chm393236
f1_keywords:
- vbapb10.chm393236
ms.prod: publisher
api_name:
- Publisher.Page.XOffsetWithinReaderSpread
ms.assetid: 42ae7545-78f5-c034-33b4-f8c8f6a0b935
ms.date: 06/08/2017
---


# Page.XOffsetWithinReaderSpread Property (Publisher)

Returns a  **Single** that represents the distance (in points) from the left edge of the reader spread to the left edge of the page. Read-only.


## Syntax

 _expression_. **XOffsetWithinReaderSpread**

 _expression_A variable that represents a  **Page** object.


### Return Value

Single


## Example

This example creates a shape on the second and third pages of the active publication and then sets the position of the shape on the third page to the diagonally opposite corner of the page from the shape on the second page. For this example to work, the active publication must have at least three pages.


```vb
Sub OffsetShapePositions() 
 Dim shpOne As Shape 
 Dim intLeft As Integer 
 Dim intTop As Integer 
 Dim intWidth As Integer 
 Dim intHeight As Integer 
 
 With ActiveDocument 
 .ViewTwoPageSpread = True 
 
 With .Pages 
 intWidth = 150 
 intHeight = 150 
 intLeft = (.Item(2).Width / 2) - intWidth 
 intTop = InchesToPoints(7) 
 
 Set shpOne = .Item(2).Shapes.AddShape _ 
 (Type:=msoShape5pointStar, Left:=intLeft, _ 
 Top:=intTop, Width:=intWidth, Height:=intHeight) 
 
 intLeft = (.Item(3).XOffsetWithinReaderSpread - _ 
 .Item(2).XOffsetWithinReaderSpread) + (.Item(2) _ 
 .Width - shpOne.Left - shpOne.Width) 
 intTop = (.Item(3).YOffsetWithinReaderSpread - _ 
 .Item(2).YOffsetWithinReaderSpread) + (.Item(2) _ 
 .Height - shpOne.Top - shpOne.Height) 
 
 .Item(2).Shapes.AddShape Type:=msoShape5pointStar, _ 
 Left:=intLeft, Top:=intTop, Width:=intWidth, _ 
 Height:=intHeight 
 End With 
 End With 
End Sub
```


