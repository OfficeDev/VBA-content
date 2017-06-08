---
title: Shape.ClassID Property (Visio)
keywords: vis_sdr.chm11213240
f1_keywords:
- vis_sdr.chm11213240
ms.prod: visio
api_name:
- Visio.Shape.ClassID
ms.assetid: b3cb2f9c-1247-9799-69f3-5374a112af95
ms.date: 06/08/2017
---


# Shape.ClassID Property (Visio)

Returns the class ID string of a shape that represents an ActiveX control or an embedded or linked OLE object. Read-only.


## Syntax

 _expression_ . **ClassID**

 _expression_ A variable that represents a **Shape** object.


### Return Value

String


## Remarks

The  **ClassID** property raises an exception if the shape doesn't represent an ActiveX control or an OLE 2.0 embedded or linked object. A shape represents an ActiveX control or an OLE 2.0 embedded or linked object if the **visTypeIsOLE2** bit (&;H8000) is set in the value returned by **Shape** . **ForeignType** .

 **ClassID** returns a string of the form:




```
{2287DC42-B167-11CE-88E9-002AFDDD917}
```

This identifies the application that services the object. It might, for example, identify an embedded object on a Microsoft Visio page as a Microsoft Excel object.

After using a shape's  **Object** property to obtain an Automation interface on the object the shape represents, you might want to obtain the shape's **ClassID** or **ProgID** property to determine the methods and properties provided by the interface.


