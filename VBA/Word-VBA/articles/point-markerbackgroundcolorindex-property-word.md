---
title: Point.MarkerBackgroundColorIndex Property (Word)
keywords: vbawd10.chm262144074
f1_keywords:
- vbawd10.chm262144074
ms.prod: word
api_name:
- Word.Point.MarkerBackgroundColorIndex
ms.assetid: 13e3de72-9090-4bd1-3dfa-51f89818ea32
ms.date: 06/08/2017
---


# Point.MarkerBackgroundColorIndex Property (Word)

Returns or sets the marker background color as an index into the current color palette, or as one of the following  **[XlColorIndex](xlcolorindex-enumeration-word.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone** . Read/write **Long** .


## Syntax

 _expression_ . **MarkerBackgroundColorIndex**

 _expression_ A variable that represents a **[Point](point-object-word.md)** object.


## Remarks

The property applies only to line, scatter, and radar charts. 


## Example

The following example sets the marker background and foreground colors for the second point in series one for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.SeriesCollection(1).Points(2) 
 ' Set the background color to green. 
 .MarkerBackgroundColorIndex = 4 
 
 ' Set the foreground color to red. 
 .MarkerForegroundColorIndex = 3 
 End With 
 End If 
End With 

```


## See also


#### Concepts


[Point Object](point-object-word.md)

