---
title: ChartGroup.BubbleScale Property (Word)
keywords: vbawd10.chm263454756
f1_keywords:
- vbawd10.chm263454756
ms.prod: word
api_name:
- Word.ChartGroup.BubbleScale
ms.assetid: 4776723c-4d6e-1009-8d00-6f837fbd4803
ms.date: 06/08/2017
---


# ChartGroup.BubbleScale Property (Word)

Returns or sets the scale factor for bubbles in the specified chart group. Read/write  **Long** .


## Syntax

 _expression_ . **BubbleScale**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-word.md)** object.


## Remarks

You can set this property to an integer from 0 (zero) through 300, corresponding to a percentage of the default size. 


 **Note**  This property applies only to bubble charts.


## Example

The following example sets the bubble size in the first chart group of the first chart in the active document to 200 percent of the default size.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.ChartGroups(1).BubbleScale = 200 
 End If 
End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-word.md)

