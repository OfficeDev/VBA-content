---
title: TaskItem.ReminderSoundFile Property (Outlook)
keywords: vbaol11.chm1739
f1_keywords:
- vbaol11.chm1739
ms.prod: outlook
api_name:
- Outlook.TaskItem.ReminderSoundFile
ms.assetid: 29bfa689-08b6-f963-9ecb-3744b1032062
ms.date: 06/08/2017
---


# TaskItem.ReminderSoundFile Property (Outlook)

Returns or sets a  **String** indicating the path and file name of the sound file to play when the reminder occurs for the Outlook item. Read/write.


## Syntax

 _expression_ . **ReminderSoundFile**

 _expression_ A variable that represents a **TaskItem** object.


## Remarks

This property is only valid if the  **[ReminderOverrideDefault](taskitem-reminderoverridedefault-property-outlook.md)** and **[ReminderPlaySound](taskitem-reminderplaysound-property-outlook.md)** properties are set to **True** .


## See also


#### Concepts


[TaskItem Object](taskitem-object-outlook.md)

