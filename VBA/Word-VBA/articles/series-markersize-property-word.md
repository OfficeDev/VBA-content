---
title: Series.MarkerSize Property (Word)
keywords: vbawd10.chm123732199
f1_keywords:
- vbawd10.chm123732199
ms.prod: word
api_name:
- Word.Series.MarkerSize
ms.assetid: fbda4404-b067-94fe-4202-a736a246e949
ms.date: 06/08/2017
---


# Series.MarkerSize Property (Word)

Returns or sets the data-marker size, in points. Read/write  **Long** .


## Syntax

 _expression_ . **MarkerSize**

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


## Remarks

This property can have a value from 2 through 72. 


## Example

The following example sets the data-marker size for all data markers on series one for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).MarkerSize = 10 
 End If 
End With 

```


## See also


#### Concepts


[Series Object](series-object-word.md)

