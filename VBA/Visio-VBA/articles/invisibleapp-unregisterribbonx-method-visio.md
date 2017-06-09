---
title: InvisibleApp.UnregisterRibbonX Method (Visio)
keywords: vis_sdr.chm17562095
f1_keywords:
- vis_sdr.chm17562095
ms.prod: visio
api_name:
- Visio.InvisibleApp.UnregisterRibbonX
ms.assetid: e32ca983-df29-0062-eb44-a5a54f334485
ms.date: 06/08/2017
---


# InvisibleApp.UnregisterRibbonX Method (Visio)

Unregisters a previouly registered  **IRibbonExtensiblity** interface that a Microsoft Visio add-in implements.


## Syntax

 _expression_ . **UnregisterRibbonX**( **_SourceAddOn_** , **_TargetDocument_** )

 _expression_ A variable that represents an **[InvisibleApp](invisibleapp-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SourceAddOn_|Required| **IRibbonExtensibility**|The add-in to unregister.|
| _TargetDocument_|Required| **[Document](document-object-visio.md)**|The document in which to unregister the add-in.|

### Return Value

 **HRESULT**


