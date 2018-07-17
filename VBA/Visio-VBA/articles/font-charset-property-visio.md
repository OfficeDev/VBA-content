---
title: Font.CharSet Property (Visio)
keywords: vis_sdr.chm12013235
f1_keywords:
- vis_sdr.chm12013235
ms.prod: visio
api_name:
- Visio.Font.CharSet
ms.assetid: 2658818f-0678-a8c2-cd4c-3628a6158a01
ms.date: 06/08/2017
---


# Font.CharSet Property (Visio)

Returns the Microsoft Windows character set for a  **Font** object. Read-only.


## Syntax

 _expression_ . **CharSet**

 _expression_ A variable that represents a **Font** object.


### Return Value

Integer


## Remarks

The Windows character set specifies character mapping for a font. The possible values of the  **CharSet** property correspond to those of the **lfCharSet** member of the Windows **LOGFONT** data structure. For details, search for " **LOGFONT** " in the Microsoft Platform SDK on MSDN, the Microsoft Developer Network.


