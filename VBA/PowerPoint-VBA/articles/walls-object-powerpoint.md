---
title: Walls Object (PowerPoint)
keywords: vbapp10.chm723000
f1_keywords:
- vbapp10.chm723000
ms.prod: powerpoint
api_name:
- PowerPoint.Walls
ms.assetid: b2288a5f-efec-84b4-9a40-d62d61196ac8
ms.date: 06/08/2017
---


# Walls Object (PowerPoint)

Represents the walls of a 3-D chart. 


## Remarks

This object is not a collection. There is no object that represents a single wall; you must return all the walls as a unit.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[Walls](chart-walls-property-powerpoint.md)** property to return the **Walls** object. The following example sets the pattern on the walls for the first chart in the active document. If the chart is not a 3-D chart, this example will fail.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Walls.Interior.Pattern = xlGray75

    End If

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

