---
title: Style.ID Property (Visio)
keywords: vis_sdr.chm11413675
f1_keywords:
- vis_sdr.chm11413675
ms.prod: visio
api_name:
- Visio.Style.ID
ms.assetid: 0eb9f8ce-302e-6749-544e-cde95fe80c72
ms.date: 06/08/2017
---


# Style.ID Property (Visio)

Gets the ID of an object. Read-only.


## Syntax

 _expression_ . **ID**

 _expression_ A variable that represents a **Style** object.


### Return Value

Long


## Remarks

The ID of a shape is unique only within the scope of the page or master. The ID of a page, master, or style is unique within the scope of the document.

If a shape, page, master, or style is deleted, future objects in the same scope may be assigned the same ID. Therefore, persisting shape or style IDs in separate data stores is generally not as sound as persisting unique IDs using the  **UniqueID** property.


