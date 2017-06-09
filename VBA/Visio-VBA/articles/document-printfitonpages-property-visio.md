---
title: Document.PrintFitOnPages Property (Visio)
keywords: vis_sdr.chm10514115
f1_keywords:
- vis_sdr.chm10514115
ms.prod: visio
api_name:
- Visio.Document.PrintFitOnPages
ms.assetid: d129ad36-0728-b3b5-60b5-3ba52e102cc7
ms.date: 06/08/2017
---


# Document.PrintFitOnPages Property (Visio)

Indicates whether drawings in a document are printed on a specified number of sheets across and down. Read/write.


## Syntax

 _expression_ . **PrintFitOnPages**

 _expression_ A variable that represents a **Document** object.


### Return Value

Boolean


## Remarks

The  **PrintFitOnPages** property corresponds to the **Fit to** settings in the **Page Setup** dialog box (on the **Design** tab, click the arrow in the **Page Setup** group). If this property is **True** , Microsoft Visio prints the document's drawings on the number of sheets specified by the **PrintPagesAcross** and **PrintPagesDown** properties.


