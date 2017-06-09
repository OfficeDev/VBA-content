---
title: ApplicationSettings.DefaultSaveFormat Property (Visio)
keywords: vis_sdr.chm16251830
f1_keywords:
- vis_sdr.chm16251830
ms.prod: visio
api_name:
- Visio.ApplicationSettings.DefaultSaveFormat
ms.assetid: 892953a8-1e69-000a-3099-c6f4baa69079
ms.date: 06/08/2017
---


# ApplicationSettings.DefaultSaveFormat Property (Visio)

Determines the default format for saving Microsoft Visio files. Read/write.


## Syntax

 _expression_ . **DefaultSaveFormat**

 _expression_ A variable that represents an **ApplicationSettings** object.


### Return Value

VisDefaultSaveFormats


## Remarks

Setting the  **DefaultSaveFormat** property is equivalent to setting the **Save files in this format** option on the **Save** tab in the **Visio Options** dialog box (click the **File** tab, click **Options**, and then click  **Save**).


 **Note**  The  **DefaultSaveFormat** property setting has no effect on the file type in which Visio files are saved by the **Save** or **SaveAs** methods of the **Document** object. To control the file type in which a document is saved programatically, use the **Version** property of the **Document** object.

The following  **VisDefaultSaveFormats** constants, which are declared in the Visio type libary, show the possible values for the **DefaultSaveFormat** property.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visDefaultSaveCurrentBinary**|0|Binary format for current version of Visio.|
| **visDefaultSaveCurrentXML**|2|XML format for current version of Visio.|
| **visDefaultSavePreviousBinary**|1|Binary format for previous version of Visio.|

