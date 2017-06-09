---
title: MailItem.ReminderSoundFile Property (Outlook)
keywords: vbaol11.chm1351
f1_keywords:
- vbaol11.chm1351
ms.prod: outlook
api_name:
- Outlook.MailItem.ReminderSoundFile
ms.assetid: 11c5ae79-1ce0-5890-1ba1-5a39a88ecc6b
ms.date: 06/08/2017
---


# MailItem.ReminderSoundFile Property (Outlook)

Returns or sets a  **String** indicating the path and file name of the sound file to play when the reminder occurs for the Outlook item. Read/write.


## Syntax

 _expression_ . **ReminderSoundFile**

 _expression_ A variable that represents a **MailItem** object.


## Remarks

This property is only valid if the  **[ReminderOverrideDefault](mailitem-reminderoverridedefault-property-outlook.md)** and **[ReminderPlaySound](mailitem-reminderplaysound-property-outlook.md)** properties are set to **True** .


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

