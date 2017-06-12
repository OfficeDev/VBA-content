---
title: Chart.AutoScaling Property (Word)
keywords: vbawd10.chm79364159
f1_keywords:
- vbawd10.chm79364159
ms.prod: word
api_name:
- Word.Chart.AutoScaling
ms.assetid: 911bf146-f3fa-7c05-a0eb-9e2062ed4a93
ms.date: 06/08/2017
---


# Chart.AutoScaling Property (Word)

 **True** if Microsoft Word scales a 3-D chart so that it is closer in size to the equivalent 2-D chart. The **[RightAngleAxes](chart-rightangleaxes-property-word.md)** property must be **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **AutoScaling**

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


## Example

The following example automatically scales the first chart in the active document. The example should be run on a 3-D chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.RightAngleAxes = True 
 .Chart.AutoScaling = True 
 End If 
End With
```


## See also


#### Concepts


[Chart Object](chart-object-word.md)

