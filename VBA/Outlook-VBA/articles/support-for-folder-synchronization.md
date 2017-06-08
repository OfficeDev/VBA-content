---
title: Support for Folder Synchronization
keywords: vbaol11.chm5272702
f1_keywords:
- vbaol11.chm5272702
ms.prod: outlook
ms.assetid: d1f941dd-fde5-b547-0751-79d03144c6bb
ms.date: 06/08/2017
---


# Support for Folder Synchronization

Users who travel with their computers or who otherwise need to use Microsoft Outlook when disconnected from the network need to be able to synchronize their offline folders using different criteria, depending on the situation. For example, before departing on a trip, a user might synchronize all of her offline folders, plus the Address Book. When she arrives at her destination, she connects to her home office using a modem. Because of the slow data-transfer rate, she only wants to synchronize her Inbox and Outbox to receive and send messages.

Outlook supports multiple synchronization profiles so users can select how they want Outlook to synchronize offline folders in a given situation. The  [SyncObjects](syncobjects-object-outlook.md) collection object represents all the synchronization profiles for a given user. Your program can use the  [Start](syncobject-start-method-outlook.md) and  [Stop](syncobject-stop-method-outlook.md) methods of the  [SyncObject](syncobject-object-outlook.md) objects in this collection to begin and end synchronization using a particular profile, and can monitor the progress of the synchronization using the  [SyncStart](syncobject-syncstart-event-outlook.md),  [Progress](syncobject-progress-event-outlook.md),  [OnError](syncobject-onerror-event-outlook.md), and  [SyncEnd Event](syncobject-syncend-event-outlook.md) events.

