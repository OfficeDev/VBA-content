---
title: Series.XValues Property (Word)
keywords: vbawd10.chm123733079
f1_keywords:
- vbawd10.chm123733079
ms.prod: word
api_name:
- Word.Series.XValues
ms.assetid: 4f558f99-dc9a-a979-9c21-a9b625716cce
ms.date: 06/08/2017
---


# Series.XValues Property (Word)

Returns or sets an array of x values for a chart series. Read/write  **Variant** .


## Syntax

 _expression_ . **XValues**

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


## Remarks

You can set the  **XValues** property to a range on a worksheet or to an array of values, but not to a combination of both.

For PivotChart reports, this property is read-only.


## Example

The following example sets the values from a range address.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).XValues = "=Sheet1!B1:B5" 
 End If 
End With
```

To assign a constant value to each individual data point, you must use an array, as shown in the following example.




```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).XValues = _ 
 Array(5.0, 6.3, 12.6, 28, 50) 
 End If 
End With
```


## See also


#### Concepts


[Series Object](series-object-word.md)

