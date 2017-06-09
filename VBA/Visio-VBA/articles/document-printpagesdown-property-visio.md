---
title: Document.PrintPagesDown Property (Visio)
keywords: vis_sdr.chm10514130
f1_keywords:
- vis_sdr.chm10514130
ms.prod: visio
api_name:
- Visio.Document.PrintPagesDown
ms.assetid: eacf72d7-f784-7a2b-0579-8af7991430ea
ms.date: 06/08/2017
---


# Document.PrintPagesDown Property (Visio)

Gets or sets the number of sheets of paper on which a drawing is printed vertically. Read/write.


## Syntax

 _expression_ . **PrintPagesDown**

 _expression_ A variable that represents a **Document** object.


### Return Value

Integer


## Remarks

You must set the value of the  **PrintFitOnPages** property to **True** to use the **PrintPagesDown** property. If the value of the **PrintFitOnPages** property is **False** , Microsoft Visio ignores the **PrintPagesDown** property.

The  **PrintPagesDown** property corresponds to the **Fit by sheet(s) down** setting in the **Page Setup** dialog box (on the **Design** tab, click the arrow in the **Page Setup** group).


