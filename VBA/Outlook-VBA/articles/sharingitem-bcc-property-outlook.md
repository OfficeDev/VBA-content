---
title: SharingItem.BCC Property (Outlook)
keywords: vbaol11.chm633
f1_keywords:
- vbaol11.chm633
ms.prod: outlook
api_name:
- Outlook.SharingItem.BCC
ms.assetid: e13c7fab-5ce6-289a-35d0-ffea5d0bd09e
ms.date: 06/08/2017
---


# SharingItem.BCC Property (Outlook)

Returns a  **String** representing the display list of blind carbon copy (BCC) names for a **[SharingItem](sharingitem-object-outlook.md)** . Read/write.


## Syntax

 _expression_ . **BCC**

 _expression_ A variable that represents a **SharingItem** object.


## Remarks

This property contains only the display names, delimited with semicolon (;) characters. The  **[Recipients](recipients-object-outlook.md)** collection should be used to modify the BCC recipients.


 **Note**  If the  **SharingItem** uses an Exchange sharing context, then setting this property to any value other than **Nothing** prevents the item from being sent and causes the **[Send](sharingitem-send-method-outlook.md)** method to raise an error.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

