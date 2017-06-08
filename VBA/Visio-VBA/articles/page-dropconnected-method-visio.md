---
title: Page.DropConnected Method (Visio)
keywords: vis_sdr.chm10962125
f1_keywords:
- vis_sdr.chm10962125
ms.prod: visio
api_name:
- Visio.Page.DropConnected
ms.assetid: 7e16dc46-df74-4482-91a4-b0a115f979b2
ms.date: 06/08/2017
---


# Page.DropConnected Method (Visio)

Creates a new  **[Shape](shape-object-visio.md)** object on the page, places the new shape relative to the specified existing target shape, and adds a connector from the existing shape to the new shape. Returns the newly created shape.


## Syntax

 _expression_ . **DropConnected**( **_ObjectToDrop_** , **_TargetShape_** , **_PlacementDir_** , **_[Connector]_** )

 _expression_ A variable that represents a **[Page](page-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectToDrop_|Required| **[UNKNOWN]**|The shape to add to the page. Can be a  **[Master](master-object-visio.md)** , **[MasterShortcut](mastershortcut-object-visio.md)** , **Shape** , or an **IDataObject** object.|
| _TargetShape_|Required| **Shape**|The existing shape from which to align, space, and connect.|
| _PlacementDir_|Required| **[VisAutoConnectDir](visautoconnectdir-enumeration-visio.md)**|The direction from  _TargetShape_ in which to place _ObjectToDrop_.|
| _Connector_|Optional| **[UNKNOWN]**|The connector to use. Can be a  **Master** , **MasterShortcut** , **Shape** , or an **IDataObject** object.|

### Return Value

 **Shape**


## Remarks

The  _ObjectToDrop_ parameter must be an object that references a two-dimensional (2-D) shape. If you pass a selection of shapes represented by an **IDataObject** object, Visio uses only the first of those shapes. If _ObjectToDrop_ is not a valid Visio object, Visio returns an Invalid Parameter error. If _ObjectToDrop_ is not a shape that matches the context of the method, Visio returns an Invalid Source error.

The  _TargetShape_ parameter must be a 2-D top-level shape on the page. If _TargetShape_ is invalid, Visio returns an Invalid Source error.

The  _PlacementDir_ parameter value must be one of the **VisAutoConnectDir** constants. If you pass **visAutoConnectDirNone** for _PlacementDir_ , Visio places the shape in a default location (0,0) and then connects it; the shape is not placed in relation to the target.

The  _Connector_ parameter must be an object that references a one-dimensional (1-D) routable shape. If you pass a selection of shapes represented by an **IDataObject** object, Visio uses only the first of those shapes. If _Connector_ is not a valid Visio object, Visio returns an Invalid Parameter error. If _Connector_ is not a shape that matches the context of the method, Visio returns an Invalid Source error.


