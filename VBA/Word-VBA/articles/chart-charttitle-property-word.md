---
title: Chart.ChartTitle Property (Word)
keywords: vbawd10.chm79364099
f1_keywords:
- vbawd10.chm79364099
ms.prod: word
api_name:
- Word.Chart.ChartTitle
ms.assetid: 1804d06a-bb2b-5995-7750-2ada70ddd1d4
ms.date: 06/08/2017
---


# Chart.ChartTitle Property (Word)

Returns the title of the specified chart. Read-only  **[ChartTitle](charttitle-object-word.md)** .


## Syntax

 _expression_ . **ChartTitle**

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


## Remarks

The  **ChartTitle** object does not exist and cannot be used unless the **[HasTitle](chart-hastitle-property-word.md)** property for the chart is **True** .


## Example

The following example sets the text for the title of the first chart.


```vb
With ActiveDocument.InlineShapes(1).Chart 
 .HasTitle = True 
 .ChartTitle.Text = "First Quarter Sales" 
End With
```


## See also


#### Concepts


[Chart Object](chart-object-word.md)

