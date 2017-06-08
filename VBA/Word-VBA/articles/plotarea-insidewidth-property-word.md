---
title: PlotArea.InsideWidth Property (Word)
keywords: vbawd10.chm53479045
f1_keywords:
- vbawd10.chm53479045
ms.prod: word
api_name:
- Word.PlotArea.InsideWidth
ms.assetid: acdb721d-73a9-15af-d833-d044e83b3c87
ms.date: 06/08/2017
---


# PlotArea.InsideWidth Property (Word)

Returns or sets the inside width, in points, of the plot area. Read/write  **Double** .


## Syntax

 _expression_ . **InsideWidth**

 _expression_ A variable that represents a **[PlotArea](plotarea-object-word.md)** object.


## Remarks

The plot area used for this measurement does not include the axis labels. The  **[Width](plotarea-width-property-word.md)** property for the plot area uses the bounding rectangle that includes the axis labels.


## Example

The following example draws a dotted rectangle around the inside of the plot area for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart 
 Set pa = .PlotArea 
 With .Shapes.AddShape(msoShapeRectangle, _ 
 pa.InsideLeft, pa.InsideTop, _ 
 pa.InsideWidth, pa.InsideHeight) 
 .Fill.Transparency = 1 
 .Line.DashStyle = msoLineDashDot 
 End With 
 End With 
 End If 
End With
```


## See also


#### Concepts


[PlotArea Object](plotarea-object-word.md)

