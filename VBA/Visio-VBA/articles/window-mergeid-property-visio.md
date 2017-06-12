---
title: Window.MergeID Property (Visio)
keywords: vis_sdr.chm11650720
f1_keywords:
- vis_sdr.chm11650720
ms.prod: visio
api_name:
- Visio.Window.MergeID
ms.assetid: 473baaa6-ea88-46f3-3d5f-501f280792a3
ms.date: 06/08/2017
---


# Window.MergeID Property (Visio)

Specifies the string version of a merged window's globally unique identifier (GUID). Read/write.


## Syntax

 _expression_ . **MergeID**

 _expression_ A variable that represents a **Window** object.


### Return Value

String


## Remarks

If this  **Window** object is not merged, the GUID will contain all zeros (GUID_NULL).

The  **MergeID** property applies only to anchored windows. If the **Window** object is an MDI frame window, Microsoft Visio raises an exception.

Use the  **Type** property to determine window type.


