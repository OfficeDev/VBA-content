---
title: Page.ID Property (Visio)
keywords: vis_sdr.chm10913675
f1_keywords:
- vis_sdr.chm10913675
ms.prod: visio
api_name:
- Visio.Page.ID
ms.assetid: 61904830-7949-98c0-eb69-a6d685b3a38c
ms.date: 06/08/2017
---


# Page.ID Property (Visio)

Gets the ID of an object. Read-only.


## Syntax

 _expression_ . **ID**

 _expression_ A variable that represents a **Page** object.


### Return Value

Long


## Remarks

The ID of a shape is unique only within the scope of the page or master. The ID of a page, master, or style is unique within the scope of the document.

If a shape, page, master, or style is deleted, future objects in the same scope may be assigned the same ID. Therefore, persisting shape or style IDs in separate data stores is generally not as sound as persisting unique IDs using the  **UniqueID** property.


