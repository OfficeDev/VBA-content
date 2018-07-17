---
title: Chart.DepthPercent Property (Word)
keywords: vbawd10.chm79364100
f1_keywords:
- vbawd10.chm79364100
ms.prod: word
api_name:
- Word.Chart.DepthPercent
ms.assetid: fd1a83dc-e68d-82be-d2bf-5f7a87cb08ac
ms.date: 06/08/2017
---


# Chart.DepthPercent Property (Word)

Returns or sets the depth of a 3-D chart as a percentage of the chart width (between 20 and 2000 percent). Read/write  **Long** .


## Syntax

 _expression_ . **DepthPercent**

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


## Remarks

This property applies only to 3-D charts.


## Example

The following example sets the depth of the first chart in the active document to be 50 percent of its width. You should run this example on a 3-D chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 Chart.DepthPercent = 50 
 End If 
End With 

```


## See also


#### Concepts


[Chart Object](chart-object-word.md)

