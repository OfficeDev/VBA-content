---
title: OlSharingProvider Enumeration (Outlook)
keywords: vbaol11.chm3146
f1_keywords:
- vbaol11.chm3146
ms.prod: outlook
api_name:
- Outlook.OlSharingProvider
ms.assetid: b42c20dc-e90d-264b-38d7-686cd74a547f
ms.date: 06/08/2017
---


# OlSharingProvider Enumeration (Outlook)

Indicates the sharing provider associated with a  **[SharingItem](sharingitem-object-outlook.md)** object.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olProviderExchange**|1|Represents the Exchange sharing provider.|
| **olProviderFederate**|7|Represents a federated sharing provider. A  **SharingItem** object with this type of provider is used for sharing relationships across organizational boundares (for example, between two organizations using Microsoft Exchange Server 2010).|
| **olProviderICal**|4|Represents the iCalendar sharing provider.|
| **olProviderPubCal**|3|Represents the PubCal sharing provider.|
| **olProviderRSS**|6|Represents the Really Simple Syndication (RSS) sharing provider.|
| **olProviderSharePoint**|5|Represents the SharePoint sharing provider.|
| **olProviderUnknown**|0|Represents an unknown sharing provider. This value is used if the sharing provider GUID in the sharing message does not match the GUID of any of the sharing providers represented in this enumeration.|
| **olProviderWebCal**|2|Represents the WebCal sharing provider.|

