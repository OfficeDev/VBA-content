---
title: PbReplaceTint Enumeration (Publisher)
keywords: vbapb10.chm65629
f1_keywords:
- vbapb10.chm65629
ms.prod: publisher
api_name:
- Publisher.PbReplaceTint
ms.assetid: 91950561-224f-285e-6dee-7d2bdbd3a589
ms.date: 06/08/2017
---


# PbReplaceTint Enumeration (Publisher)

Constants passed to the  **Plate.Delete** method specifying how to replace the colors in a deleted plate.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **pbReplaceTintKeepTints**|1|Maintain the same tint percentage in the ink represented by the replacement plate as in the deleted plate. For example, replace a 100% tint of yellow with a 100% tint of blue.|
| **pbReplaceTintMaintainLuminosity**|2|Maintain the same lightness value in the ink represented by the replacement plate as in the deleted plate. For example, replace a 100% tint of yellow with an approximately 10% tint of blue.|
| **pbReplaceTintUseDefault**|0|Default.|

