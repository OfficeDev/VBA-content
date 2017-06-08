---
title: DataLabel.ShowRange Property (Word)
keywords: vbawd10.chm233900019
f1_keywords:
- vbawd10.chm233900019
ms.prod: word
ms.assetid: c9e3e8e5-630e-cb5b-ed48-5842dee505e9
ms.date: 06/08/2017
---


# DataLabel.ShowRange Property (Word)

Set to  **True** to display the **Value From Cells** range field for the specified chart data label. Set to **False** to hide that field. Read/write **Boolean**.


## Syntax

 _expression_ . **ShowRange**

 _expression_ A variable that represents a **DataLabel** object.


## Example

The following example displays the  **Value From Cells** range field for the first data label of the first series on the first chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).DataLabels(1).ShowRange = True 
 End If 
End With
```


## Property value

 **BOOL**


## See also


#### Concepts


[DataLabel Object](datalabel-object-word.md)

