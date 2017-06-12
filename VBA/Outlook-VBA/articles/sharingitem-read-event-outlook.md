---
title: SharingItem.Read Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.SharingItem.Read
ms.assetid: 2bcf07e6-e9c1-b3ce-118c-a2c82b48ff5f
ms.date: 06/08/2017
---


# SharingItem.Read Event (Outlook)

Occurs when an instance of the parent object is opened for editing by the user. 


## Syntax

 _expression_ . **Read**

 _expression_ An expression that returns a **SharingItem** object.


## Remarks

The  **Read** event differs from the **[Open](sharingitem-open-event-outlook.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an **[Inspector](inspector-object-outlook.md)** .


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

