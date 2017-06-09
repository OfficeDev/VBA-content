---
title: Reminders.Snooze Event (Outlook)
keywords: vbaol11.chm580
f1_keywords:
- vbaol11.chm580
ms.prod: outlook
api_name:
- Outlook.Reminders.Snooze
ms.assetid: 253e3f16-6d33-e7f7-5a1f-4a8b0a82a55d
ms.date: 06/08/2017
---


# Reminders.Snooze Event (Outlook)

Occurs when a reminder is dismissed using the  **Snooze** button.


## Syntax

 _expression_ . **Snooze**( **_ReminderObject_** )

 _expression_ An expression that returns a **Reminders** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ReminderObject_|Required| **[Reminder](reminder-object-outlook.md)**|Represents the reminder to dismiss.|

## Remarks

This event will fire when the  **[Snooze](reminder-snooze-method-outlook.md)** method is executed, or when the user clicks the **Snooze** button.


## Example

The following Microsoft Visual Basic for Applications (VBA) example displays the original date and time set for the  **Reminder** object that has been snoozed.


```vb
Public WithEvents objReminders As Outlook.Reminders 
 
Sub Initialize_Handler() 
 Set objReminders = Application.Reminders 
End Sub 
 
Private Sub objReminders_Snooze(ByVal ReminderObject As Reminder) 
 'Occurs when a user clicks Snooze or when snooze is 
 'programmatically executed. 
 MsgBox "The reminder was originally set at " _ 
 &; ReminderObject.OriginalReminderDate 
End Sub
```


## See also


#### Concepts


[Reminders Object](reminders-object-outlook.md)

