---
title: Series.MarkerStyle Property (Word)
keywords: vbawd10.chm123732040
f1_keywords:
- vbawd10.chm123732040
ms.prod: word
api_name:
- Word.Series.MarkerStyle
ms.assetid: d9ba7847-2785-0f29-7e6e-d4bb2d62fc2f
ms.date: 06/08/2017
---


# Series.MarkerStyle Property (Word)

Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write  **[XlMarkerStyle](xlmarkerstyle-enumeration-word.md)** .


## Syntax

 _expression_ . **MarkerStyle**

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


## Remarks

 **MarkerStyle** can be one of the following **XlMarkerStyle** constants:


-  **xlMarkerStyleAutomatic** —Automatic markers.
    
-  **xlMarkerStyleCircle** —Circular markers.
    
-  **xlMarkerStyleDash** —Long bar markers.
    
-  **xlMarkerStyleDiamond** —Diamond-shaped markers.
    
-  **xlMarkerStyleDot** —Short bar markers.
    
-  **xlMarkerStyleNone** —No markers.
    
-  **xlMarkerStylePicture** —Picture markers.
    
-  **xlMarkerStylePlus** —Square markers with a plus sign.
    
-  **xlMarkerStyleSquare** —Square markers.
    
-  **xlMarkerStyleStar** —Square markers with an asterisk.
    
-  **xlMarkerStyleTriangle** —Triangular markers.
    
-  **xlMarkerStyleX** —Square markers with an X.
    



## Example

The following example sets the marker style for series one for the first chart in the active document. You should run the example on a 2-D line chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle 
 End If 
End With
```


## See also


#### Concepts


[Series Object](series-object-word.md)

