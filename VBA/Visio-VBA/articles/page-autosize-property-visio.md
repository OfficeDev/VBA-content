---
title: Page.AutoSize Property (Visio)
keywords: vis_sdr.chm10962450
f1_keywords:
- vis_sdr.chm10962450
ms.prod: visio
ms.assetid: 777155fb-21a6-f7d2-3eef-66ed09a00628
ms.date: 06/08/2017
---


# Page.AutoSize Property (Visio)

Determines whether Microsoft Visio automatically resizes the drawing page by adding printer-paper-sized sheets, as necessary, to fit the drawing's contents. Read/write.


## Syntax

 _expression_ . **AutoSize**

 _expression_ An expression that returns a **[Page](page-object-visio.md)** object.


### Return Value

 **Boolean**


## Remarks

Set  **AutoSize** to **True** to enable automatic resizing of the page. Set **AutoSize** to **False** (the default) to disable automatic resizing.

The  **AutoSize** property setting corresponds to the state of the **AutoSize** button in the **Page Setup** group on the **Design** tab.


