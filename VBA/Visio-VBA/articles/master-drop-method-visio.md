---
title: Master.Drop Method (Visio)
keywords: vis_sdr.chm10716235
f1_keywords:
- vis_sdr.chm10716235
ms.prod: visio
api_name:
- Visio.Master.Drop
ms.assetid: 13abc8fc-7b3c-98cf-3965-3ac7b3d15e85
ms.date: 06/08/2017
---


# Master.Drop Method (Visio)

Creates one or more new  **Shape** objects by dropping an object onto a receiving object such as a master, drawing page, shape, or group.


## Syntax

 _expression_ . **Drop**( **_ObjectToDrop_** , **_xPos_** , **_yPos_** )

 _expression_ A variable that represents a **Master** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectToDrop_|Required| **[UNKNOWN]**|The object or selection to drop. While this is typically a Visio object such as a  **Master** , **Shape** , or **Selection** object, it can be any OLE object that provides an **IDataObject** interface.|
| _xPos_|Required| **Double**|The x-coordinate at which to place the center of the shape's width or PinX.|
| _yPos_|Required| **Double**|The y-coordinate at which to place the center of the shape's height or PinY.|

### Return Value

Shape


## Remarks

Using the  **Drop** method is similar to moving a shape with the mouse. The object dropped ( _ObjectToDrop_) can be a master or a shape on the drawing page.

To add a shape to a group or on a drawing page, apply the  **Drop** method to a **Shape** or **Page** object, respectively. The center of the shape's width-height box is positioned at the specified coordinates, and a **Shape** object that represents the shape that is created is returned. When applying this method to a **Shape** object, make sure that the **Shape** object represents a group.

If  _ObjectToDrop_ is a **Master** , the pin of the master is dropped at the specified coordinates. A master's pin is often, but not necessarily, at its center of rotation.


