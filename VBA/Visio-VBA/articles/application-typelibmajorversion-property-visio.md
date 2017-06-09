---
title: Application.TypelibMajorVersion Property (Visio)
keywords: vis_sdr.chm10014695
f1_keywords:
- vis_sdr.chm10014695
ms.prod: visio
api_name:
- Visio.Application.TypelibMajorVersion
ms.assetid: 17e1abf3-5a5d-aac9-9f78-4eeb2c4a6c79
ms.date: 06/08/2017
---


# Application.TypelibMajorVersion Property (Visio)

Returns the major version number of the Microsoft Visio type library. Read-only.


## Syntax

 _expression_ . **TypelibMajorVersion**

 _expression_ A variable that represents an **Application** object.


### Return Value

Integer


## Remarks

The major and/or minor version number of the Visio type library will increase whenever the Visio type library is extended. A program can use the  **TypelibMajorVersion** and **TypelibMinorVersion** properties to guarantee that the Visio version it is working with provides support for the features it is using.

Small changes to the Visio type library do not affect the  **Application** object's **Version** property.


