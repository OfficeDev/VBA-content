---
title: Series.Values Property (Word)
keywords: vbawd10.chm123732132
f1_keywords:
- vbawd10.chm123732132
ms.prod: word
api_name:
- Word.Series.Values
ms.assetid: 5dc3ebb2-9da6-59c5-c388-78dbf88551df
ms.date: 06/08/2017
---


# Series.Values Property (Word)

Returns or sets a collection of all the values in the series. Read/write  **Variant** .


## Syntax

 _expression_ . **Values**

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


## Remarks

The value of this property can be either the address of a range on the chart's worksheet or an array of constant values, but not a combination of both. See the examples for details.


## Example

The following example sets the series values from a range address.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Values = "=Sheet1!B2:B5" 
 End If 
End With
```

To assign a constant value to each individual data point, you must use an array, as shown in the following example.




```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Values = _ 
 Array(1, 3, 5, 7, 11, 13, 17, 19) 
 End If 
End With
```


## See also


#### Concepts


[Series Object](series-object-word.md)

