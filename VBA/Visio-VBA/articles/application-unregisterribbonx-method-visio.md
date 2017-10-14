---
title: Application.UnregisterRibbonX Method (Visio)
keywords: vis_sdr.chm10062095
f1_keywords:
- vis_sdr.chm10062095
ms.prod: visio
api_name:
- Visio.Application.UnregisterRibbonX
ms.assetid: 83c5a7c3-b3bb-cbbd-6857-2ae822e599f6
ms.date: 06/08/2017
---


# Application.UnregisterRibbonX Method (Visio)

Unregisters a previouly registered  **IRibbonExtensiblity** interface that a Microsoft Visio add-in implements.


## Syntax

 _expression_ . **UnregisterRibbonX**( **_SourceAddOn_** , **_TargetDocument_** )

 _expression_ A variable that represents an **[Application](application-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SourceAddOn_|Required| **IRibbonExtensibility**|The add-in to unregister.|
| _TargetDocument_|Required| **[Document](document-object-visio.md)**|The document in which to unregister the add-in.|

### Return Value

 **Nothing**


