---
title: Application.TypelibMinorVersion Property (Visio)
keywords: vis_sdr.chm10014700
f1_keywords:
- vis_sdr.chm10014700
ms.prod: visio
api_name:
- Visio.Application.TypelibMinorVersion
ms.assetid: ee3a31db-ddfe-a036-a570-43e6f27ad024
ms.date: 06/08/2017
---


# Application.TypelibMinorVersion Property (Visio)

Returns the minor version number of the Microsoft Visio type library. Read-only.


## Syntax

 _expression_ . **TypelibMinorVersion**

 _expression_ A variable that represents an **Application** object.


### Return Value

Integer


## Remarks

The major and/or minor version number of the Visio type library will increase whenever the Visio type library is extended. A program can use the  **TypelibMajorVersion** and **TypelibMinorVersion** properties to guarantee that the Visio version it is working with provides support for the features it is using.

Small changes to the Visio type library do not affect the  **Application** object's **Version** property.


