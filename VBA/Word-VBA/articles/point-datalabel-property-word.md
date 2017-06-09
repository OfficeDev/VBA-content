---
title: Point.DataLabel Property (Word)
keywords: vbawd10.chm262144158
f1_keywords:
- vbawd10.chm262144158
ms.prod: word
api_name:
- Word.Point.DataLabel
ms.assetid: d84afe14-7c11-8ccf-baf0-687b72f25314
ms.date: 06/08/2017
---


# Point.DataLabel Property (Word)

Returns the data label associated with the point. Read-only  **[DataLabel](datalabel-object-word.md)** .


## Syntax

 _expression_ . **DataLabel**

 _expression_ A variable that represents a **[Point](point-object-word.md)** object.


## Example

The following example enables the data label for point seven in series three of the first chart in the active document, and then it sets the data label color to blue.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.SeriesCollection(3).Points(7) 
 .HasDataLabel = True 
 .ApplyDataLabels type:=xlValue 
 .DataLabel.Font.ColorIndex = 5 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Point Object](point-object-word.md)

