---
title: Chart.Rotation Property (Word)
keywords: vbawd10.chm79364112
f1_keywords:
- vbawd10.chm79364112
ms.prod: word
api_name:
- Word.Chart.Rotation
ms.assetid: a141124f-f33c-95e1-6ba9-8ecffdef434c
ms.date: 06/08/2017
---


# Chart.Rotation Property (Word)

Returns or sets the rotation, in degrees, of the 3-D chart view (the rotation of the plot area around the z-axis). Read/write  **Variant** .


## Syntax

 _expression_ . **Rotation**

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


## Remarks

The value of this property must be from 0 through 360, except for 3-D bar charts, where the value must be from 0 through 44. The default value is 20. This property applies only to 3-D charts. 

Rotations are always rounded to the nearest integer.


## Example

The following example sets the rotation of the first chart in the active document to 30 degrees. You should run the example on a 3-D chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Rotation = 30 
 End If 
End With
```


## See also


#### Concepts


[Chart Object](chart-object-word.md)

