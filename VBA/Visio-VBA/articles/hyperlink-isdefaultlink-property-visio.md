---
title: Hyperlink.IsDefaultLink Property (Visio)
keywords: vis_sdr.chm15013720
f1_keywords:
- vis_sdr.chm15013720
ms.prod: visio
api_name:
- Visio.Hyperlink.IsDefaultLink
ms.assetid: 5a958e11-cf88-c45d-829a-805af9fd9f3a
ms.date: 06/08/2017
---


# Hyperlink.IsDefaultLink Property (Visio)

Determines the default  **Hyperlink** object for a shape. Read/write.


## Syntax

 _expression_ . **IsDefaultLink**

 _expression_ A variable that represents a **Hyperlink** object.


### Return Value

Integer


## Remarks

When you set the value of the  **IsDefaultLink** property to **True** for a **Hyperlink** object, the value for all other **Hyperlink** objects is automatically set to **False** . When you set the value of this property to **False** for a **Hyperlink** object, the other **Hyperlink** objects aren't affected.


