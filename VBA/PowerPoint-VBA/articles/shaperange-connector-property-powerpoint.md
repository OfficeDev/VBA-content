---
title: ShapeRange.Connector Property (PowerPoint)
keywords: vbapp10.chm548020
f1_keywords:
- vbapp10.chm548020
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Connector
ms.assetid: 04871183-d9d0-f0ba-f801-4f1f6564f0d3
ms.date: 06/08/2017
---


# ShapeRange.Connector Property (PowerPoint)

Determines whether the specified shape is a connector. Read-only.


## Syntax

 _expression_. **Connector**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

MsoTriState


## Remarks

The value of the  **Connector** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified shape is not a connector.|
|**msoTrue**| The specified shape is a connector.|

## Example

This example deletes all connectors on  `myDocument`.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes

    For i = .Count To 1 Step -1

        With .Item(i)

            If .Connector Then .Delete

        End With

    Next

End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

