---
title: Borders.DistanceFromBottom Property (Word)
keywords: vbawd10.chm154927125
f1_keywords:
- vbawd10.chm154927125
ms.prod: word
api_name:
- Word.Borders.DistanceFromBottom
ms.assetid: 97184500-0536-33ed-1552-80ea829f0e30
ms.date: 06/08/2017
---


# Borders.DistanceFromBottom Property (Word)

Returns or sets the space (in points) between the text and the bottom border. Read/write  **Long** .


## Syntax

 _expression_ . **DistanceFromBottom**

 _expression_ A variable that represents a **[Borders](borders-object-word.md)** object.


## Remarks

Using this property with a page border, you can set either the space between the text and the bottom page border or the space between the bottom edge of the page and the bottom page border. Where the distance is measured from depends on the value of the  **[DistanceFrom](borders-distancefrom-property-word.md)** property.


## Example

This example adds a border around the first paragraph in the active document and sets the distance between the text and the bottom border to 6 points.


```vb
With ActiveDocument.Paragraphs(1).Borders 
 .Enable = True 
 .DistanceFromBottom = 6 
End With
```

This example adds a border around each table in Sales.doc. The example also sets the distance between the text and the border to 3 points for the top and bottom borders, and 6 points for the left and right borders.




```vb
Dim tableLoop As Table 
 
For Each tableLoop In Documents("Sales.doc").Tables 
 With tableLoop.Borders 
 .OutsideLineStyle = wdLineStyleSingle 
 .OutsideLineWidth = wdLineWidth150pt 
 .DistanceFromBottom = 3 
 .DistanceFromTop = 3 
 .DistanceFromLeft = 6 
 .DistanceFromRight = 6 
 End With 
Next tableLoop
```


## See also


#### Concepts


[Borders Collection Object](borders-object-word.md)

