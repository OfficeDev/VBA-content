---
title: Series.HasErrorBars Property (Word)
keywords: vbawd10.chm123732128
f1_keywords:
- vbawd10.chm123732128
ms.prod: word
api_name:
- Word.Series.HasErrorBars
ms.assetid: c41f951a-c483-249e-1384-02b6180d5835
ms.date: 06/08/2017
---


# Series.HasErrorBars Property (Word)

 **True** if the series has error bars. Read/write **Boolean** .


## Syntax

 _expression_ . **HasErrorBars**

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


## Remarks

This property is not available for 3-D charts. 


## Example

The following example removes error bars from series one for the first chart in the active document. You should run the example on a 2-D line chart that has error bars for series one.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).HasErrorBars = False 
 End If 
End With
```


## See also


#### Concepts


[Series Object](series-object-word.md)

