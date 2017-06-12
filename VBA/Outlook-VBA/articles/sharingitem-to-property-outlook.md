---
title: SharingItem.To Property (Outlook)
keywords: vbaol11.chm665
f1_keywords:
- vbaol11.chm665
ms.prod: outlook
api_name:
- Outlook.SharingItem.To
ms.assetid: a9a7d504-9baf-7408-4b4b-240123aebaa8
ms.date: 06/08/2017
---


# SharingItem.To Property (Outlook)

Returns or sets a semicolon-delimited  **String** list of display names for the **To** recipients for the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.


## Syntax

 _expression_ . **To**

 _expression_ A variable that represents a **SharingItem** object.


## Remarks

This property contains the display names only. The  **To** property corresponds to the MAPI property **PidTagDisplayTo** . The **[Recipients](recipients-object-outlook.md)** collection should be used to modify this property.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

