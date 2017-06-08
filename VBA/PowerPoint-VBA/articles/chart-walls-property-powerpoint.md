---
title: Chart.Walls Property (PowerPoint)
keywords: vbapp10.chm684047
f1_keywords:
- vbapp10.chm684047
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.Walls
ms.assetid: e4c019c0-41de-988b-b5c7-009fcc0eee15
ms.date: 06/08/2017
---


# Chart.Walls Property (PowerPoint)

Returns the walls of the 3-D chart. Read-only  **[Walls](walls-object-powerpoint.md)**.


## Syntax

 _expression_. **Walls**

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the color of the wall border of the first chart in the active document to red. You should run the example on a 3-D chart.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.Walls.Border. _
            ColorIndex = 3
    End If
End With


```


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

