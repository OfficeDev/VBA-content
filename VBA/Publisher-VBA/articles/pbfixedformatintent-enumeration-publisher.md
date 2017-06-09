---
title: PbFixedFormatIntent Enumeration (Publisher)
keywords: vbapb10.chm65637
f1_keywords:
- vbapb10.chm65637
ms.prod: publisher
api_name:
- Publisher.PbFixedFormatIntent
ms.assetid: bddb023b-181f-7805-434f-128f27d609e4
ms.date: 06/08/2017
---


# PbFixedFormatIntent Enumeration (Publisher)

Constants passed to the  **[ExportAsFixedFormat](document-exportasfixedformat-method-publisher.md)** method that specify how the user intends to share the resulting file.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **pbIntentCommercial**|4|Submit the publication to a commercial press.|
| **pbIntentMinimum**|1|Squeeze the publication to the smallest file size. This satisfies the on-screen viewing scenario where the publication is viewed on a computer monitor.|
| **pbIntentPrinting**|3|Print the publication on a desktop printer or at a copy store, such as Kinko's.|
| **pbIntentStandard**|2|Distribute the publication as an e-mail message or from a Web site. Note that the user does not know how the publication will be viewed: on-screen or printed from a desktop printer. Both the desktop printing scenario and the on-screen viewing scenario must be met by this intent.|

