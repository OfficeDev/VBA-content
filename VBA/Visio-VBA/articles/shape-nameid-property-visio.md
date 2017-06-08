---
title: Shape.NameID Property (Visio)
keywords: vis_sdr.chm11213935
f1_keywords:
- vis_sdr.chm11213935
ms.prod: visio
api_name:
- Visio.Shape.NameID
ms.assetid: ae658ed9-124f-22f2-53be-5c9b6ebaa382
ms.date: 06/08/2017
---


# Shape.NameID Property (Visio)

Returns a unique name for a shape. Read-only.


## Syntax

 _expression_ . **NameID**

 _expression_ A variable that represents a **Shape** object.


### Return Value

String


## Remarks

The  **NameID** property returns a unique identifier for each shape on a page or master. The identifier has the following form: sheet. _N_ , where _N_ is the shape's **ID** property. The value of the **NameID** property is unique within a page or master, but not across pages or masters. At any moment, no other shape on the same page or master has the same **NameID** property. However, shapes on other pages or masters may have the same **NameID** property.

The value of a shape's  **UniqueID** property is unique across pages and masters.

Also,  **NameID** properties are reused. If a shape whose **NameID** property is sheet. _N_ is deleted, a shape subsequently added to the same context may be assigned sheet.N as its **NameID** property. Therefore, persisting **NameID** properties in separate data stores is generally not as sound as persisting **UniqueID** properties.


