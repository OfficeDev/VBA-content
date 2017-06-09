---
title: Shape.Disconnect Method (Visio)
keywords: vis_sdr.chm11262255
f1_keywords:
- vis_sdr.chm11262255
ms.prod: visio
api_name:
- Visio.Shape.Disconnect
ms.assetid: ece61baa-dfe7-7b61-5c45-49de4cf0e394
ms.date: 06/08/2017
---


# Shape.Disconnect Method (Visio)

Unglues the specified connector end points and offsets them the specified amount from the shapes to which they were joined.


## Syntax

 _expression_ . **Disconnect**( **_ConnectorEnd_** , **_OffsetX_** , **_OffsetY_** , **_Units_** )

 _expression_ A variable that represents a **[Shape](shape-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ConnectorEnd_|Required| **[VisConnectorEnds](visconnectorends-enumeration-visio.md)**|The end of the connector to disconnect.|
| _OffsetX_|Required| **Double**|The  _x_-distance that the connector end is moved away from the shape.|
| _OffsetY_|Required| **Double**|The  _y_-distance that the connector end is moved away from the shape.|
| _Units_|Required| **[VisUnitCodes](visunitcodes-enumeration-visio.md)**|The units of measure for the assigned offset values.|

### Return Value

 **Nothing**


## Remarks

 _ConnectorEnd_ must be one of the following **VisConnectorEnds** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visConnectorBeginPoint**|0|The begin point of the connector.|
| **visConnectorEndPoint**|1|The end point of the connector.|
| **visConnectorBothEnds**|2|Both the begin and the end point of the connector.|
When you call  **Disconnect** on a connector shape (a 1-D routable shape), one or both endpoints of the connector are unglued from their target shapes, based on the specified _ConnectorEnd_ parameter value. If a specified endpoint is not glued, Microsoft Visio takes no action.

Visio offsets the endpoint(s) from their current position by the amount specified by  _OffsetX_ , _OffsetY_ , and _Units_ . Offset values of 0 mean that the endpoints do not move.

The  **Disconnect** method does not apply to non-connector shapes. If you call **Disconnect** on a non-connector shape or on a shape in a master, Visio returns an Invalid Source error.


