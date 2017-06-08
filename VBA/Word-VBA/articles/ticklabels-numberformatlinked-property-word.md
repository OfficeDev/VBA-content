---
title: TickLabels.NumberFormatLinked Property (Word)
keywords: vbawd10.chm167051270
f1_keywords:
- vbawd10.chm167051270
ms.prod: word
api_name:
- Word.TickLabels.NumberFormatLinked
ms.assetid: c0daa894-b45e-69c1-540a-fa91599b105b
ms.date: 06/08/2017
---


# TickLabels.NumberFormatLinked Property (Word)

 **True** if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells). Read/write **Boolean** .


## Syntax

 _expression_ . **NumberFormatLinked**

 _expression_ A variable that represents a **[TickLabels](ticklabels-object-word.md)** object.


## Example

The following example links the number format for tick-mark labels to its cells for the value axis of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Axes(xlValue).TickLabels.NumberFormatLinked = True 
 End If 
End With
```


## See also


#### Concepts


[TickLabels Object](ticklabels-object-word.md)

