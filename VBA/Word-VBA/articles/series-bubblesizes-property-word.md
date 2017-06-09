---
title: Series.BubbleSizes Property (Word)
keywords: vbawd10.chm123733632
f1_keywords:
- vbawd10.chm123733632
ms.prod: word
api_name:
- Word.Series.BubbleSizes
ms.assetid: 4972506e-9a19-417d-0ef3-a259b68a7ed1
ms.date: 06/08/2017
---


# Series.BubbleSizes Property (Word)

Returns or sets a string that refers to the worksheet cells that contain the x-value, y-value, and size data for the bubble chart. Read/write  **Variant** .


## Syntax

 _expression_ . **BubbleSizes**

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


## Remarks

 When you return the cell reference, it will return a string that describes the cells in A1-style notation. To set the size data for the bubble chart, you must use R1C1-style notation.


 **Note**  This property applies only to bubble charts.


## Example

The following example displays the cell reference for the cells that contain the bubble chart x-value, y-value, and size data for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 MsgBox .Chart.SeriesCollection(1).BubbleSizes 
 End If 
End With
```

This example shows how to set this property using R1C1-style notation.




```vb
With ActiveDocument.InlineShapes(3) 
 If .HasChart Then 
 .Chart.SeriesCollection(1). _ 
 BubbleSizes = "Sheet1!r2c3:r5c3" 
 End If 
End With
```


## See also


#### Concepts


[Series Object](series-object-word.md)

