---
title: Chart.Floor Property (PowerPoint)
keywords: vbapp10.chm684029
f1_keywords:
- vbapp10.chm684029
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.Floor
ms.assetid: a599ed02-d137-f823-0a5e-325d3f8c7406
ms.date: 06/08/2017
---


# Chart.Floor Property (PowerPoint)

Returns the floor of the 3-D chart. Read-only  **[Floor](floor-object-powerpoint.md)**.


## Syntax

 _expression_. **Floor**

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the floor color of the first chart in the active document to blue. You should run the example on a 3-D chart (the  **Floor** property fails on 2-D charts).




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Floor.Interior.ColorIndex = 5

    End If

End With


```


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

