---
title: MeetingItem.SendUsingAccount Property (Outlook)
keywords: vbaol11.chm3509
f1_keywords:
- vbaol11.chm3509
ms.prod: outlook
api_name:
- Outlook.MeetingItem.SendUsingAccount
ms.assetid: 81713c7b-dfb0-eb91-b017-82b427bee823
ms.date: 06/08/2017
---


# MeetingItem.SendUsingAccount Property (Outlook)

Returns or sets an  **[Account](account-object-outlook.md)** object that represents the account to use to send the **[MeetingItem](meetingitem-object-outlook.md)** . Read/write.


## Syntax

 _expression_ . **SendUsingAccount**

 _expression_ A variable that represents a **MeetingItem** object.


## Remarks

You can use the  **SendUsingAccount** property to specify the account that the **Send** method uses to send the **MeetingItem** . This property returns **Null** ( **Nothing** in Visual Basic) if the account specified for the **MeetingItem** no longer exists.

This property is read-only if the  **MeetingItem** is a received item, or if the **MeetingItem** has already been sent (that is, the **[Sent](meetingitem-sent-property-outlook.md)** property of the object is set to **True** ).


## See also


#### Concepts


[MeetingItem Object](meetingitem-object-outlook.md)

