---
title: OlAttachmentBlockLevel Enumeration (Outlook)
keywords: vbaol11.chm3261
f1_keywords:
- vbaol11.chm3261
ms.prod: outlook
api_name:
- Outlook.OlAttachmentBlockLevel
ms.assetid: 651fced7-9853-255e-66ed-7aa5f52c1b9c
ms.date: 06/08/2017
---


# OlAttachmentBlockLevel Enumeration (Outlook)

Specifies whether there is any restriction on the type of attachments for an item.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olAttachmentBlockLevelNone**|0|There is no restriction on the type of the attachment based on its file extension.|
| **olAttachmentBlockLevelOpen**|1|There is a restriction on the type of the attachment based on its file extension such that users must first save the attachment to disk before opening it.|

## Remarks

Attachments with the [BlockLevel](attachment-blocklevel-property-outlook.md) equal to **olAttachmentBlockLevelOpen** are on the Level 2 list of attachments that administrators maintain for attachment security. For more information on attachment security in Outlook, see the[Office Resource Kit](http://technet.microsoft.com/en-us/library/cc303401%28office.14%29.aspx) Web site.


