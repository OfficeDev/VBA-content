---
title: InvisibleApp.AutoLayout Property (Visio)
keywords: vis_sdr.chm17513105
f1_keywords:
- vis_sdr.chm17513105
ms.prod: visio
api_name:
- Visio.InvisibleApp.AutoLayout
ms.assetid: 46f2a65d-a86c-9750-8879-69081187b061
ms.date: 06/08/2017
---


# InvisibleApp.AutoLayout Property (Visio)

Allows you to temporarily disable automatic layout functionality in Microsoft Visio and then re-enable it after you are finished with an action. Read/write.


## Syntax

 _expression_ . **AutoLayout**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

Boolean


## Remarks

Using the  **AutoLayout** property helps to improve the performance of add-ons that execute many operations in connected drawings that use Visio automatic layout functionality.


