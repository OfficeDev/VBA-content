---
title: ChartGroup.HasHiLoLines Property (Word)
keywords: vbawd10.chm263454732
f1_keywords:
- vbawd10.chm263454732
ms.prod: word
api_name:
- Word.ChartGroup.HasHiLoLines
ms.assetid: 5713e885-9f36-6b6c-2622-a813cba2077b
ms.date: 06/08/2017
---


# ChartGroup.HasHiLoLines Property (Word)

 **True** if the line chart has high-low lines. Read/write **Boolean** .


## Syntax

 _expression_ . **HasHiLoLines**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-word.md)** object.


## Remarks

This property applies only to line charts. 


## Example

The following example enables high-low lines for chart group one of the first chart in the active document and then sets line style, weight, and color. You should run the example on a 2-D line chart that has three series of stock-quote-like data (high-low-close).


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.ChartGroups(1) 
 .HasHiLoLines = True 
 With .HiLoLines.Border 
 .LineStyle = xlThin 
 .Weight = xlMedium 
 .ColorIndex = 3 
 End With 
 End With 
 End If 
End With 

```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-word.md)

