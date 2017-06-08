---
title: OLEObject.Object Property (Visio)
keywords: vis_sdr.chm15213950
f1_keywords:
- vis_sdr.chm15213950
ms.prod: visio
api_name:
- Visio.OLEObject.Object
ms.assetid: 802c5651-96d0-511b-a403-7e43019bb051
ms.date: 06/08/2017
---


# OLEObject.Object Property (Visio)

Returns an  **IDispatch** interface on the ActiveX control or embedded or linked OLE 2.0 object represented by a **Shape** object or an **OLEObject** object. Read-only.


## Syntax

 _expression_ . **Object**

 _expression_ A variable that represents an **OLEObject** object.


### Return Value

Object


## Remarks

The  **Object** property raises an exception if the object doesn't represent an ActiveX control or an OLE 2.0 embedded or linked object. A shape represents an ActiveX control or an OLE 2.0 embedded or linked object if the **visTypeIsOLE2** bit (&;H8000) is set in the value returned by the **ForeignType** property.

If the  **Object** property succeeds, it returns an **IDispatch** interface on the control or object. You owe an eventual release on the returned value (set it to **Nothing** or let it go out of scope if you're using Microsoft Visual Basic). You can determine the kind of object you've obtained an interface on by using the **ClassID** or **ProgID** property.

Beginning with Microsoft Visio 5.0, if the object returned by the  **Object** property is embedded and the shape inherits the object from its master, the **Object** property severs the instance—that is, it copies the inherited data into the instance. Otherwise, if the client receiving the **IDispatch** interface from the **Object** property makes changes to the object, all instances of the master, not just the instance being queried, change. If the object returned by the **Object** property is linked, the **Object** property does not sever the instance because, by definition, there may be other entities referencing the link. The **ObjectIsInherited** property was added to Visio 5.0 so that client programs can know if a shape inherits its object and access the master's object(s).


