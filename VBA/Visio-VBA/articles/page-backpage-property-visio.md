---
title: Page.BackPage Property (Visio)
keywords: vis_sdr.chm10913115
f1_keywords:
- vis_sdr.chm10913115
ms.prod: visio
api_name:
- Visio.Page.BackPage
ms.assetid: cef2dac4-cf12-d692-cbbc-a6023f2d78e0
ms.date: 06/08/2017
---


# Page.BackPage Property (Visio)

Gets or sets the background page of a page. Read/write.


## Syntax

 _expression_ . **BackPage**

 _expression_ A variable that represents a **Page** object.


### Return Value

Variant


## Remarks

If a page has no background, its  **BackPage** property returns an empty **Variant** . Otherwise the returned **Variant** refers to a **Page** object, the background page of the indicated page.

To assign a background page to a page, set the page's  **BackPage** property to the name of that background page. To cause a page to have no background page, pass an empty string to the **BackPage** property.

Markup overlay pages cannot have background pages, so you cannot use the  **BackPage** property to assign a background page to a markup overlay page.


 **Note**  In earlier versions of Visio (through version 4.1), the  **BackPage** property returned an object (as opposed to a **Variant** of type **Object** ) and it accepted a string (as opposed to a **Variant** of type **String** ). Because of changes in Automation support tools, the property has been modified so that it accepts and returns a **Variant** .


