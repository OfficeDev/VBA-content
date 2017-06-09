---
title: UIObject.DisplayKeysInTooltips Property (Visio)
keywords: vis_sdr.chm14913410
f1_keywords:
- vis_sdr.chm14913410
ms.prod: visio
api_name:
- Visio.UIObject.DisplayKeysInTooltips
ms.assetid: 81cd5ba8-decb-0db7-391d-b79bfbdc4cb6
ms.date: 06/08/2017
---


# UIObject.DisplayKeysInTooltips Property (Visio)

Determines whether ScreenTip text includes keyboard shortcuts. Read/write. 


## Syntax

 _expression_ . **DisplayKeysInTooltips**

 _expression_ A variable that represents a **UIObject** object.


### Return Value

Boolean


## Remarks

To show ScreenTips, you must set the  **DisplayTooltips** property to **True** .

It does not matter which  **UIObject** object you use when getting or setting this property. The property affects the entire application.

This property setting corresponds to the  **Show shortcut keys in ScreenTips** setting on the **General** tab in the **Visio Options** dialog box (click the **File** tab, and then click **Options**), and is shared between Visio and all Microsoft Office applications.


