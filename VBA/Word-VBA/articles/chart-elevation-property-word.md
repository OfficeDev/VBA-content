---
title: Chart.Elevation Property (Word)
keywords: vbawd10.chm79364102
f1_keywords:
- vbawd10.chm79364102
ms.prod: word
api_name:
- Word.Chart.Elevation
ms.assetid: 2913dce4-e35a-31ff-3ea0-237c85dbad23
ms.date: 06/08/2017
---


# Chart.Elevation Property (Word)

Returns or sets the elevation, in degrees, of the 3-D chart view. Read/write  **Long** .


## Syntax

 _expression_ . **Elevation**

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


## Remarks

The chart elevation is the height, in degrees, at which you view the chart. The default is 15 for most chart types. The value of this property must be between -90 and 90, except for 3-D bar charts, where it must be between 0 and 44.


 **Note**  This property applies only to 3-D charts.


## Example

The following example sets the chart elevation of the first chart in the active document to 34 degrees. You should run the example on a 3-D chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Elevation = 34 
 End If 
End With
```


## See also


#### Concepts


[Chart Object](chart-object-word.md)

