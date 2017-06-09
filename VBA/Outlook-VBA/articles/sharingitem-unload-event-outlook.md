---
title: SharingItem.Unload Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.SharingItem.Unload
ms.assetid: b79a4c94-46cc-5571-a36d-ad537db97bcc
ms.date: 06/08/2017
---


# SharingItem.Unload Event (Outlook)

Occurs before an Outlook item is unloaded from memory, either programmatically or by user action. 


## Syntax

 _expression_ . **Unload**

 _expression_ An expression that returns a **SharingItem** object.


## Remarks

This event occurs after the  **Close** event for the Outlook item occurs, but before the Outlook item is unloaded from memory, allowing an add-in to release any resources related to the object. Although the event occurs before the Outlook item is unloaded from memory, this event cannot be canceled.


 **Note**  This event is meant only as a notification event, so that an add-in can dereference the object. An error occurs if any property or method for this object is called within the  **Unload** event.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

