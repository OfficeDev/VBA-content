---
title: MasterShortcut.ExportIcon Method (Visio)
keywords: vis_sdr.chm16016270
f1_keywords:
- vis_sdr.chm16016270
ms.prod: visio
api_name:
- Visio.MasterShortcut.ExportIcon
ms.assetid: 763f6dac-98fb-377d-484b-08bac9b55d90
ms.date: 06/08/2017
---


# MasterShortcut.ExportIcon Method (Visio)

Exports the icon for a  **Master** object to a named file or the Clipboard.


## Syntax

 _expression_ . **ExportIcon**( **_FileName_** , **_Flags_** , **_TransparentRGB_** )

 _expression_ A variable that represents a **MasterShortcut** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The file to which to export the icon.|
| _Flags_|Required| **Integer**|The format in which to write the exported file.|
| _TransparentRGB_|Optional| **Variant**|The color to substitute for any transparent areas of the exported icon image.|

### Return Value

Nothing


## Remarks

If  _FileName_ is empty, the master's icon is copied to the Clipboard.

If the value of  _Flags_ is **visIconFormatVisio** (0), the icon is exported in the application internal icon format. The **ImportIcon** method accepts files written in this format.

If the value of  _Flags_ is **visIconFormatBMP** (2), the icon is exported in bitmap (.bmp) file format.

Starting with Microsoft Visio 2000, you can use the  _TransparentRGB_ argument with the **ExportIcon** method. If _TransparentRGB_ is omitted, the color defaults to black, which simulates Visio 5.0 behavior.


