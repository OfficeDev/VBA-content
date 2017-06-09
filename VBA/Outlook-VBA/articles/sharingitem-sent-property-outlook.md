---
title: SharingItem.Sent Property (Outlook)
keywords: vbaol11.chm661
f1_keywords:
- vbaol11.chm661
ms.prod: outlook
api_name:
- Outlook.SharingItem.Sent
ms.assetid: 6ae38f11-186e-3c86-f131-4eb6230f10a7
ms.date: 06/08/2017
---


# SharingItem.Sent Property (Outlook)

Returns a  **Boolean** value that indicates if the **[SharingItem](sharingitem-object-outlook.md)** has been sent. Read-only.


## Syntax

 _expression_ . **Sent**

 _expression_ A variable that represents a **SharingItem** object.


## Remarks

In general, there are three different kinds of messages: sent, posted, and saved. Sent messages are items sent to a recipient or public folder. Posted messages are created in a public folder. Saved messages are created and saved without either sending or posting.

If this property is set to  **True** , attempting to set the **[SendUsingAccount](sharingitem-sendusingaccount-property-outlook.md)** property raises an error.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

