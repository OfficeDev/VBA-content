---
title: Document.PrintPagesAcross Property (Visio)
keywords: vis_sdr.chm10514125
f1_keywords:
- vis_sdr.chm10514125
ms.prod: visio
api_name:
- Visio.Document.PrintPagesAcross
ms.assetid: 43c09ce5-fcc9-d91c-3108-5e2dcb658f74
ms.date: 06/08/2017
---


# Document.PrintPagesAcross Property (Visio)

Indicates the number of sheets of paper on which a drawing is printed horizontally. Read/write.


## Syntax

 _expression_ . **PrintPagesAcross**

 _expression_ A variable that represents a **Document** object.


### Return Value

Integer


## Remarks

You must set the value of the  **PrintFitOnPages** property to **True** in order to use the **PrintPagesAcross** property. If the value of the **PrintFitOnPages** property is **False** , Microsoft Visio ignores the **PrintPagesAcross** property.

The  **PrintPagesAcross** property corresponds to the **Fit to sheet(s) across** setting in the **Page Setup** dialog box (on the **Design** tab, click the arrow in the **Page Setup** group).


