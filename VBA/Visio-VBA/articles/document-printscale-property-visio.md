---
title: Document.PrintScale Property (Visio)
keywords: vis_sdr.chm10514135
f1_keywords:
- vis_sdr.chm10514135
ms.prod: visio
api_name:
- Visio.Document.PrintScale
ms.assetid: d352b695-1e94-888d-70a0-9189678992e6
ms.date: 06/08/2017
---


# Document.PrintScale Property (Visio)

Gets or sets how much drawings are reduced or enlarged when printed. Read/write.


## Syntax

 _expression_ . **PrintScale**

 _expression_ A variable that represents a **Document** object.


### Return Value

Double


## Remarks

The  **PrintScale** property corresponds to the **Adjust to** setting on the **Print Setup** tab in the **Page Setup** dialog box (on the **Design** tab, click the arrow in the **Page Setup** group). To print a drawing at half its size, specify 0.5. To print a drawing at twice its size, specify 2.0.


