---
title: AppointmentItem.BeforeAttachmentRead Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.BeforeAttachmentRead
ms.assetid: db38a11b-c9bc-ebda-5900-00391cdf080f
ms.date: 06/08/2017
---


# AppointmentItem.BeforeAttachmentRead Event (Outlook)

Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an  **[Attachment](attachment-object-outlook.md)** object.


## Syntax

 _expression_ . **BeforeAttachmentRead**( **_Attachment_** , **_Cancel_** )

 _expression_ A variable that represents an **AppointmentItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **Attachment**|The  **Attachment** to be read.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be read.|

## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

