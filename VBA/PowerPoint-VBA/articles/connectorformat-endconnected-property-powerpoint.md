---
title: ConnectorFormat.EndConnected Property (PowerPoint)
keywords: vbapp10.chm555009
f1_keywords:
- vbapp10.chm555009
ms.prod: powerpoint
api_name:
- PowerPoint.ConnectorFormat.EndConnected
ms.assetid: b5e4b8cb-a69c-7330-5dae-0fa4b7a36c82
ms.date: 06/08/2017
---


# ConnectorFormat.EndConnected Property (PowerPoint)

Determines whether the end of the specified connector is connected to a shape. Read-only.


## Syntax

 _expression_. **EndConnected**

 _expression_ A variable that represents an **ConnectorFormat** object.


### Return Value

MsoTriState


## Remarks

The value of the  **EndConnected** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**| The end of the specified connector is not connected to a shape.|
|**msoTrue**| The end of the specified connector is connected to a shape.|

## Example

If the end of the connector represented by shape three on the first slide in the active presentation is connected to a shape, this example stores the connection site number in the variable  `oldEndConnSite`, stores a reference to the connected shape in the object variable  `oldEndConnShape`, and then disconnects the end of the connector from the shape.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3)

    If .Connector Then

        With .ConnectorFormat

            If .EndConnected Then

                oldEndConnSite = .EndConnectionSite

                Set oldEndConnShape = .EndConnectedShape

                .EndDisconnect

            End If

        End With

    End If

End With
```


## See also


#### Concepts


[ConnectorFormat Object](connectorformat-object-powerpoint.md)

