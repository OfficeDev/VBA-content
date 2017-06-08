---
title: Master.ImportIcon Method (Visio)
keywords: vis_sdr.chm10716360
f1_keywords:
- vis_sdr.chm10716360
ms.prod: visio
api_name:
- Visio.Master.ImportIcon
ms.assetid: 886d724d-9d02-ab6f-8049-80fa04f8caec
ms.date: 06/08/2017
---


# Master.ImportIcon Method (Visio)

Imports the icon for a  **Master** object from a named file.


## Syntax

 _expression_ . **ImportIcon**( **_FileName_** )

 _expression_ A variable that represents a **Master** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The name of the file to import.|

### Return Value

Nothing


## Remarks

The  **ImportIcon** method can only import files that were produced by exporting a master icon in the application's internal icon format ( **visIconFormatVisio** )?it does not accept icons in other file formats.


