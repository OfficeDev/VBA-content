---
title: TickLabels.NumberFormatLinked Property (PowerPoint)
keywords: vbapp10.chm719006
f1_keywords:
- vbapp10.chm719006
ms.prod: powerpoint
api_name:
- PowerPoint.TickLabels.NumberFormatLinked
ms.assetid: df60a8dc-85be-7e7e-68ea-0a60a60ef977
ms.date: 06/08/2017
---


# TickLabels.NumberFormatLinked Property (PowerPoint)

 **True** if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells). Read/write **Boolean**.


## Syntax

 _expression_. **NumberFormatLinked**

 _expression_ A variable that represents a **[TickLabels](ticklabels-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

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


[TickLabels Object](ticklabels-object-powerpoint.md)

