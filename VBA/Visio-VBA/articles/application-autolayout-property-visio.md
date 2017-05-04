---
title: Application.AutoLayout Property (Visio)
keywords: vis_sdr.chm10013105
f1_keywords:
- vis_sdr.chm10013105
ms.prod: VISIO
api_name:
- Visio.Application.AutoLayout
ms.assetid: b631def8-d271-8ed0-880a-db8a1ee26759
---


# Application.AutoLayout Property (Visio)

Allows you to temporarily disable automatic layout functionality in Microsoft Visio and then re-enable it after you are finished with an action. Read/write.


## Syntax

 _expression_ . **AutoLayout**

 _expression_ A variable that represents an **Application** object.


### Return Value

Boolean


## Remarks

Using the  **AutoLayout** property helps to improve the performance of add-ons that execute many operations in connected drawings that use Visio automatic layout functionality.


