---
title: OlTrackingStatus Enumeration (Outlook)
keywords: vbaol11.chm3088
f1_keywords:
- vbaol11.chm3088
ms.prod: outlook
api_name:
- Outlook.OlTrackingStatus
ms.assetid: a2253862-b1a1-6d99-81ad-1984ba615919
ms.date: 06/08/2017
---


# OlTrackingStatus Enumeration (Outlook)

Indicates the most recent tracking status change for the recipient.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olTrackingDelivered**|1|The item has been delivered to the recipient.|
| **olTrackingNone**|0|No tracking information is available for the recipient.|
| **olTrackingNotDelivered**|2|The item has not been delivered to the recipient.|
| **olTrackingNotRead**|3|The item has not been read by the recipient.|
| **olTrackingRead**|6|The item has been read by the recipient.|
| **olTrackingRecallFailure**|4|The sender of the item attempted to recall the item but was unsuccessful.|
| **olTrackingRecallSuccess**|5|The sender of the item recalled the item.|
| **olTrackingReplied**|7|The recipient replied to the item.|

## Remarks

Used by the [Recipient.TrackingStatus Property (Outlook)](recipient-trackingstatus-property-outlook.md).


