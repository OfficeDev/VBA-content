---
title: SharingItem.RetentionExpirationDate Property (Outlook)
keywords: vbaol11.chm3565
f1_keywords:
- vbaol11.chm3565
ms.prod: outlook
api_name:
- Outlook.SharingItem.RetentionExpirationDate
ms.assetid: 29a7456d-4c2d-a418-699f-3e3984d5d0a2
ms.date: 06/08/2017
---


# SharingItem.RetentionExpirationDate Property (Outlook)

Returns a  **Date** that specifies the date when the **[SharingItem](sharingitem-object-outlook.md)** object expires, after which the Messaging Records Management (MRM) Assistant will delete the item. Read-only.


## Syntax

 _expression_ . **RetentionExpirationDate**

 _expression_ A variable that represents a **SharingItem** object.


## Remarks

A retention policy is enabled and disabled by an administrator for an Exchange Server on a mailbox level. This feature is available only on an Exchange mailbox with MRM version 2.0 or later enabled.

Microsoft Outlook calculates the value of this property based on the item retention start date and the retention period, if Outlook is in cache or offline mode. The Exchange Server specifies the value if Outlook is in online mode.

 In general, the retention start date for the item is determined as follows:


- Received or sent items: the retention start date is the received date.
    
- Nonrecurring calendar items: the retention start date is the appointment end date.
    
- Recurring calendar items: the retention start date is the end date of last recurrence. If there is no end date, the item never expires.
    



## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

