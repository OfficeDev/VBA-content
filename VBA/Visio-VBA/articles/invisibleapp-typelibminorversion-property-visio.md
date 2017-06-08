---
title: InvisibleApp.TypelibMinorVersion Property (Visio)
keywords: vis_sdr.chm17514700
f1_keywords:
- vis_sdr.chm17514700
ms.prod: visio
api_name:
- Visio.InvisibleApp.TypelibMinorVersion
ms.assetid: 7564e196-4999-037f-650f-a6fa6f9e3308
ms.date: 06/08/2017
---


# InvisibleApp.TypelibMinorVersion Property (Visio)

Returns the minor version number of the Microsoft Visio type library. Read-only.


## Syntax

 _expression_ . **TypelibMinorVersion**( **_lpi2Ret_** )

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

Integer


## Remarks

The major and/or minor version number of the Visio type library will increase whenever the Visio type library is extended. A program can use the  **TypelibMajorVersion** and **TypelibMinorVersion** properties to guarantee that the Visio version it is working with provides support for the features it is using.

Small changes to the Visio type library do not affect the  **Application** object's **Version** property.


