---
title: Application.NewMailEx Event (Outlook)
keywords: vbaol11.chm438
f1_keywords:
- vbaol11.chm438
ms.prod: outlook
api_name:
- Outlook.Application.NewMailEx
ms.assetid: 3b6873a3-0ccf-0e46-1cac-0eeabb3a896b
ms.date: 06/08/2017
---


# Application.NewMailEx Event (Outlook)

Occurs when a new item is received in the Inbox.


## Syntax

 _expression_ . **NewMailEx**( **_EntryIDCollection_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EntryIDCollection_|Required| **String**|A string representing an Entry ID of an item received in the  **Inbox**.|

## Remarks

This event fires once for every received item that is processed by Microsoft Outlook. The item can be one of several different item types, for example,  **[MailItem](mailitem-object-outlook.md)** , **[MeetingItem](meetingitem-object-outlook.md)** , or **[SharingItem](sharingitem-object-outlook.md)** . The _EntryIDsCollection_ string contains the Entry ID that corresponds to that item. Note that this behavior has changed from earlier versions of the event when the _EntryIDCollection_ contained a list of comma-delimited Entry IDs of all the items received in the Inbox since the last time the event was fired.

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).

This event fires for e-mail accounts that provide notifications for received messages, such as Microsoft Exchange Server and POP3 accounts.

The  **NewMailEx** event fires when a new message arrives in the Inbox and before client rule processing occurs. You can use the Entry ID returned in the _EntryIDCollection_ array to call the **[NameSpace.GetItemFromID](namespace-getitemfromid-method-outlook.md)** method and process the item. Use this method with caution to minimize the impact on Outlook performance. However, depending on the setup on the client computer, after a new message arrives in the Inbox, processes like spam filtering and client rules that move the new message from the Inbox to another folder can occur asynchronously. You should not assume that after these events fire, you will always get a one-item increase in the number of items in the Inbox.

For users with an Exchange Server account (non-Cached Exchange Mode or Cached Exchange Mode), the event will fire only for messages that arrive at the server after Outlook has started. The event will not fire for messages that are synchronized in Cached Exchange Mode immediately after Outlook starts, nor for messages that are already on the server when Outlook starts in non-Cached Exchange Mode.

For users using Cached Exchange Mode, the event will fire in all settings, provided that Outlook is running when the message is received:  **Download Full Items**,  **Download Headers**, and  **Download Headers and then Full Items**.


## See also


#### Concepts


[Application Object](application-object-outlook.md)

