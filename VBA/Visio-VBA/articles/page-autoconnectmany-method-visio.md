---
title: Page.AutoConnectMany Method (Visio)
keywords: vis_sdr.chm10962130
f1_keywords:
- vis_sdr.chm10962130
ms.prod: visio
api_name:
- Visio.Page.AutoConnectMany
ms.assetid: 292d0f58-d753-6ef3-fd62-269fd44d003c
ms.date: 06/08/2017
---


# Page.AutoConnectMany Method (Visio)

Automatically draws multiple connections in the specified directions between the specified shapes. Returns the number of shapes connected.


## Syntax

 _expression_ . **AutoConnectMany**( **_FromShapeIDs()_** , **_ToShapeIDs()_** , **_PlacementDirs()_** , **_[Connector]_** )

 _expression_ A variable that represents a **[Page](page-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FromShapeIDs()_|Required| **Long**|An array of identifers of the shapes from which to draw a connection.|
| _ToShapeIDs()_|Required| **Long**|An array of identifers of the shapes to which to draw a connection.|
| _PlacementDirs()_|Required| **Long**|An array of  **[VisAutoConnectDir](visautoconnectdir-enumeration-visio.md)** constants that represent the directions in which to draw the connections. See Remarks for possible values.|
| _Connector_|Optional| **[UNKNOWN]**|The connector to use. Can be a  **[Master](master-object-visio.md)** , **[MasterShortcut](mastershortcut-object-visio.md)** , **[Shape](shape-object-visio.md)** , or **IDataObject** object.|

### Return Value

 **Long**


## Remarks

For the  _PlacementDirs()_ parameter, pass an array of values from the **VisAutoConnectDir** enumeration to specify the connection directions (that is, where to locate the connected shapes with respect to the primary shapes). Possible values for _PlacementDirs()_ are as follows.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|visAutoConnectDirDown|2|Connect down.|
|visAutoConnectDirLeft|3|Connect to the left.|
|visAutoConnectDirNone|0|Connect without relocating the shapes.|
|visAutoConnectDirRight|4|Connect to the right|
|visAutoConnectDirUp|1|Connect up.|
Calling the  **AutoConnectMany** method is equivalent to calling the **[Shape.AutoConnect](shape-autoconnect-method-visio.md)** method multiple times.

You can include the same shape multiple times in each array you pass as a parameter. You cannot use the  **AutoConnectMany** method to connect a shape to itself.

If a particular  **AutoConnectMany** operation fails or is invalid, Microsoft Visio skips it and processes the next item in each of the parameter arrays. **AutoConnectMany** returns the total number of items successfully processed.

If the parameter arrays do not each contain the same number of values, Visio returns an Invalid Parameter error.

The optional  _Connector_ parameter value must be an object that references a one-dimensional routable shape. If you do not pass a value for _Connector_ , Visio uses the default dynamic connector.

If you use the  **IDataObject** interface to pass a selection of shapes for _Connector_ , Visio uses only the first shape. If _Connector_ is not a Visio object, Visio returns an Invalid Parameter error. If _Connector_ is not a shape that matches the context of the method, Visio returns an Invalid Source error.


