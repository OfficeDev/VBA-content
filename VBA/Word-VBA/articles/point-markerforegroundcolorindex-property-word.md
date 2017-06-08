---
title: Point.MarkerForegroundColorIndex Property (Word)
keywords: vbawd10.chm262144076
f1_keywords:
- vbawd10.chm262144076
ms.prod: word
api_name:
- Word.Point.MarkerForegroundColorIndex
ms.assetid: 76c259a9-da33-4de1-f6c5-c0aa75ff1ff9
ms.date: 06/08/2017
---


# Point.MarkerForegroundColorIndex Property (Word)

Returns or sets the marker foreground color as an index into the current color palette, or as one of the following  **[XlColorIndex](xlcolorindex-enumeration-word.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone** . Read/write **Long** .


## Syntax

 _expression_ . **MarkerForegroundColorIndex**

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

