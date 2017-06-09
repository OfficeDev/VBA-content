---
title: Presentation.SnapToGrid Property (PowerPoint)
keywords: vbapp10.chm583061
f1_keywords:
- vbapp10.chm583061
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.SnapToGrid
ms.assetid: d0155913-cca5-c2ed-b1cc-6463a573ff49
ms.date: 06/08/2017
---


# Presentation.SnapToGrid Property (PowerPoint)

Determines whether to snap shapes to the gridlines in the specified presentation. Read/write.


## Syntax

 _expression_. **SnapToGrid**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

MsoTriState


## Remarks

The value of the  **SnapToGrid** property can be one of these **MsoTriState** constants.


||
|:-----|
|**msoFalse**|
|**msoTrue**|

## Example

This example switches snapping shapes to the gridlines in the active presentation.


```vb
Sub ToggleSnapToGrid()

    With ActivePresentation

        If .SnapToGrid = msoTrue Then

            .SnapToGrid = msoFalse

        Else

            .SnapToGrid = msoTrue

        End If

    End With

End Sub
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

