---
title: Master.ID Property (Visio)
keywords: vis_sdr.chm10713675
f1_keywords:
- vis_sdr.chm10713675
ms.prod: visio
api_name:
- Visio.Master.ID
ms.assetid: 9064e708-f939-9522-b8f7-24488d780bc0
ms.date: 06/08/2017
---


# Master.ID Property (Visio)

Gets the ID of an object. Read-only.


## Syntax

 _expression_ . **ID**

 _expression_ A variable that represents a **Master** object.


### Return Value

Long


## Remarks

The ID of a shape is unique only within the scope of the page or master. The ID of a page, master, or style is unique within the scope of the document.

If a shape, page, master, or style is deleted, future objects in the same scope may be assigned the same ID. Therefore, persisting shape or style IDs in separate data stores is generally not as sound as persisting unique IDs using the  **UniqueID** property.


