---
title: MasterShortcut.ImportIcon Method (Visio)
keywords: vis_sdr.chm16016360
f1_keywords:
- vis_sdr.chm16016360
ms.prod: visio
api_name:
- Visio.MasterShortcut.ImportIcon
ms.assetid: f48cb1ea-e0b2-ebba-39b3-da7e6be46dcb
ms.date: 06/08/2017
---


# MasterShortcut.ImportIcon Method (Visio)

Imports the icon for a  **Master** object from a named file.


## Syntax

 _expression_ . **ImportIcon**( **_FileName_** )

 _expression_ A variable that represents a **MasterShortcut** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The name of the file to import.|

### Return Value

Nothing


## Remarks

The  **ImportIcon** method can only import files that were produced by exporting a master icon in the application's internal icon format ( **visIconFormatVisio** )?it does not accept icons in other file formats.


