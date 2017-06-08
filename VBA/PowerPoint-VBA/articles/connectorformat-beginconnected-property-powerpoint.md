---
title: ConnectorFormat.BeginConnected Property (PowerPoint)
keywords: vbapp10.chm555006
f1_keywords:
- vbapp10.chm555006
ms.prod: powerpoint
api_name:
- PowerPoint.ConnectorFormat.BeginConnected
ms.assetid: c7c2c448-590c-b1b6-8dc5-9fcb44974fee
ms.date: 06/08/2017
---


# ConnectorFormat.BeginConnected Property (PowerPoint)

Determines whether the beginning of the specified connector is connected to a shape. Read/write.


## Syntax

 _expression_. **BeginConnected**

 _expression_ A variable that represents a **ConnectorFormat** object.


### Return Value

MsoTriState


## Remarks

The value of the  **BeginConnected** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**| The beginning of the specified connector is not connected to a shape.|
|**msoTrue**| The beginning of the specified connector is connected to a shape.|

## Example

If shape three on the first slide in the active presentation is a connector whose beginning is connected to a shape, this example stores the connection site number in the variable  `oldBeginConnSite`, stores a reference to the connected shape in the object variable  `oldBeginConnShape`, and then disconnects the beginning of the connector from the shape.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3)

    If .Connector Then

        With .ConnectorFormat

            If .BeginConnected Then

                oldBeginConnSite = .BeginConnectionSite

                Set oldBeginConnShape = .BeginConnectedShape

                .BeginDisconnect

            End If

        End With

    End If

End With
```


## See also


#### Concepts


[ConnectorFormat Object](connectorformat-object-powerpoint.md)

