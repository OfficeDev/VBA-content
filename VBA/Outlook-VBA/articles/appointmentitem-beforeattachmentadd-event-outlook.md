---
title: AppointmentItem.BeforeAttachmentAdd Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.BeforeAttachmentAdd
ms.assetid: 1574e5b0-b2d1-ca0a-3e1a-0c5efef0a99c
ms.date: 06/08/2017
---


# AppointmentItem.BeforeAttachmentAdd Event (Outlook)

Occurs before an attachment is added to an instance of the parent object.


## Syntax

 _expression_ . **BeforeAttachmentAdd**( **_Attachment_** , **_Cancel_** )

 _expression_ A variable that represents an **AppointmentItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](attachment-object-outlook.md)**|The  **Attachment** to be added to the item.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be added.|

## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

