---
title: Reminders.ReminderFire Event (Outlook)
keywords: vbaol11.chm578
f1_keywords:
- vbaol11.chm578
ms.prod: outlook
api_name:
- Outlook.Reminders.ReminderFire
ms.assetid: 73a3f825-8aef-95b8-00c5-74e19daed84a
ms.date: 06/08/2017
---


# Reminders.ReminderFire Event (Outlook)

Occurs before the reminder is executed.


## Syntax

 _expression_ . **ReminderFire**( **_ReminderObject_** )

 _expression_ A variable that represents a **Reminders** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ReminderObject_|Required| **[Reminder](reminder-object-outlook.md)**|The  **Reminder** object that has been executed.|

## Example

The following Microsoft Visual Basic for Applications (VBA) example displays the item that fired the  **Reminder** event every time a reminder is executed.


```vb
Public WithEvents objReminders As Outlook.Reminders 
 
 
 
Sub Initialize_handler() 
 
 Set objReminders = Application.Reminders 
 
End Sub 
 
 
 
Private Sub objReminders_ReminderFire(ByVal ReminderObject As Reminder) 
 
 'Opens the item when a reminder executes 
 
 ReminderObject.Item.Display 
 
End Sub
```


## See also


#### Concepts


[Reminders Object](reminders-object-outlook.md)

