---
title: Axis.TickLabelSpacing Property (Word)
keywords: vbawd10.chm113049651
f1_keywords:
- vbawd10.chm113049651
ms.prod: word
api_name:
- Word.Axis.TickLabelSpacing
ms.assetid: af49728e-6c42-7846-50da-127c855264bf
ms.date: 06/08/2017
---


# Axis.TickLabelSpacing Property (Word)

Returns or sets the number of categories or series between tick-mark labels. Read/write  **Long** .


## Syntax

 _expression_ . **TickLabelSpacing**

 _expression_ A variable that represents an **[Axis](axis-object-word.md)** object.


## Remarks

This property applies only to category and series axes. It can be a value from 1 through 31999. 

Tick-mark label spacing on the value axis is always calculated by Microsoft Word.


## Example

The following example sets the number of categories between tick-mark labels on the category axis of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Axes(xlCategory).TickLabelSpacing = 10 
 End If 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-word.md)

