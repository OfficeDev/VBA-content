---
title: InvisibleApp.SetPreviewEnabled Method (Visio)
keywords: vis_sdr.chm17562105
f1_keywords:
- vis_sdr.chm17562105
ms.prod: visio
api_name:
- Visio.InvisibleApp.SetPreviewEnabled
ms.assetid: c84a2f51-0019-8f3b-26b6-ecd593a4cbe5
ms.date: 06/08/2017
---


# InvisibleApp.SetPreviewEnabled Method (Visio)

Turns preview on or off for a gallery in the Microsoft Visio user interface.


## Syntax

 _expression_ . **SetPreviewEnabled**( **_GalleryName_** , **_OnOrOff_** )

 _expression_ A variable that represents an **[InvisibleApp](invisibleapp-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _GalleryName_|Required| **String**|The name of the gallery for which to set the preview status.|
| _OnOrOff_|Required| **Boolean**| **True** to turn preview on for the specified gallery; **False** to turn preview off.|

### Return Value

 **Nothing**


## Remarks

For the  _GalleryName_ parameter, you must pass the control ID for the specified gallery. You can find a list of control IDs for all Visio galleries by searching the MSDN library at http://msdn.microsoft.com/library.


