---
title: Point.MarkerBackgroundColor Property (Word)
keywords: vbawd10.chm262144073
f1_keywords:
- vbawd10.chm262144073
ms.prod: word
api_name:
- Word.Point.MarkerBackgroundColor
ms.assetid: 629e0174-4590-3531-23ae-6093e9ca77a1
ms.date: 06/08/2017
---


# Point.MarkerBackgroundColor Property (Word)

Sets the marker background color as an RGB value or returns the corresponding color index value. Read/write  **Long** .


## Syntax

 _expression_ . **MarkerBackgroundColor**

 _expression_ A variable that represents a **[Point](point-object-word.md)** object.


## Remarks

This property applies only to line, scatter, and radar charts. 


## Example

The following example sets the marker background and foreground colors for the second point in series one for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.SeriesCollection(1).Points(2) 
 ' Set the background color to green. 
 .MarkerBackgroundColor = RGB(0,255,0) 
 
 ' Set the foreground color to red. 
 .MarkerForegroundColor = RGB(255,0,0) 
 End With 
 End If 
End With 

```


## See also


#### Concepts


[Point Object](point-object-word.md)

