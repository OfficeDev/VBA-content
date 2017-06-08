---
title: Chart.RightAngleAxes Property (Word)
keywords: vbawd10.chm79364110
f1_keywords:
- vbawd10.chm79364110
ms.prod: word
api_name:
- Word.Chart.RightAngleAxes
ms.assetid: d7f01a8f-aa76-3e92-2db2-370176066145
ms.date: 06/08/2017
---


# Chart.RightAngleAxes Property (Word)

 **True** if the chart axes are at right angles, independent of chart rotation or elevation. Read/write **Boolean** .


## Syntax

 _expression_ . **RightAngleAxes**

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


## Remarks

This property applies only to 3-D line, column, and bar charts. 

If this property is set to  **True** , the **[Perspective](chart-perspective-property-word.md)** property is ignored.


## Example

The following example sets the axes for the first chart in the active document to intersect at right angles. You should run the example on a 3-D chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.RightAngleAxes = True 
 End If 
End With
```


## See also


#### Concepts


[Chart Object](chart-object-word.md)

