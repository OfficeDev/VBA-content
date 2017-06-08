---
title: Point.Paste Method (Word)
keywords: vbawd10.chm262144211
f1_keywords:
- vbawd10.chm262144211
ms.prod: word
api_name:
- Word.Point.Paste
ms.assetid: 88a215df-a271-2d09-8ffe-765fcb990163
ms.date: 06/08/2017
---


# Point.Paste Method (Word)

Pastes a picture from the Clipboard as the marker on the selected point.


## Syntax

 _expression_ . **Paste**

 _expression_ A variable that represents a **[Point](point-object-word.md)** object.


## Remarks

You can use this method on column, bar, line, or radar charts, and it sets the  **[MarkerStyle](point-markerstyle-property-word.md)** property to **xlMarkerStylePicture** .


## Example

The following example pastes a picture from the Clipboard into point one in series one for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Points(1).Paste 
 End If 
End With 

```


## See also


#### Concepts


[Point Object](point-object-word.md)

