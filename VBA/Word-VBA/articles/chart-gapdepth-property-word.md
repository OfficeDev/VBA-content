---
title: Chart.GapDepth Property (Word)
ms.prod: word
api_name:
- Word.Chart.GapDepth
ms.assetid: 09147a74-c8bb-4fc5-0389-c8f46e0be67d
ms.date: 06/08/2017
---


# Chart.GapDepth Property (Word)

Returns or sets the distance, as a percentage of the marker width, between the data series in a 3-D chart. Read/write  **Long** .


## Syntax

 _expression_ . **GapDepth**

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


## Remarks

The value of this property must be between 0 and 500. 


 **Note**  This property applies only to 3-D charts.


## Example

The following example sets the distance between the data series for the first chart in the active document to 200 percent of the marker width. You should run the example on a 3-D chart (the  **GapDepth** property fails on 2-D charts).


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.GapDepth = 200 
 End If 
End With
```


## See also


#### Concepts


[Chart Object](chart-object-word.md)

