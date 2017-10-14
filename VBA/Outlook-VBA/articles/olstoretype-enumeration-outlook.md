---
title: OlStoreType Enumeration (Outlook)
keywords: vbaol11.chm3100
f1_keywords:
- vbaol11.chm3100
ms.prod: outlook
api_name:
- Outlook.OlStoreType
ms.assetid: a23d132f-32ae-5b4d-5d9e-aa09411f4be0
ms.date: 06/08/2017
---


# OlStoreType Enumeration (Outlook)

Indicates the format in which the data file should be created.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olStoreANSI**|3|ANSI format personal folders file (.pst) compatible with all previous versions of Microsoft Outlook format.|
| **olStoreDefault**|1|Default format compatible with the mailbox mode in which Outlook runs on the Microsoft Exchange Server.|
| **olStoreUnicode**|2|Unicode format personal folders file (.pst) compatible with Microsoft Office Outlook 2003 and later.|

## Remarks

Used as a parameter to the [NameSpace.AddStoreEx Method (Outlook)](namespace-addstoreex-method-outlook.md).


