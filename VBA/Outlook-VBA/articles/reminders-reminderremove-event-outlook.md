---
title: Reminders.ReminderRemove Event (Outlook)
keywords: vbaol11.chm579
f1_keywords:
- vbaol11.chm579
ms.prod: outlook
api_name:
- Outlook.Reminders.ReminderRemove
ms.assetid: f217cd33-84c0-223b-ad4e-9ceb0f7e894c
ms.date: 06/08/2017
---


# Reminders.ReminderRemove Event (Outlook)

Occurs when a  **[Reminder](reminder-object-outlook.md)** object has been removed from the collection.


## Syntax

 _expression_ . **ReminderRemove**

 _expression_ A variable that represents a **Reminders** object.


## Remarks

A reminder can be removed from the  **Reminders** collection by any of the following means:


- The  **Reminders** collection's **[Remove](reminders-remove-method-outlook.md)** method.
    
- The  **Reminder** object's **[Dismiss](reminder-dismiss-method-outlook.md)** method.
    
- When the user clicks the  **Dismiss** button.
    
- When a user turns off a meeting reminder from within the associated item.
    
- When a user deletes an item that contains a reminder.
    

## Example

The following Microsoft Visual Basic for Applications (VBA) example displays a message to the user when a  **[Reminder](reminder-object-outlook.md)** object is removed from the collection.


```vb
Public WithEvents objReminders As Outlook.Reminders 
 
 
 
Sub Initialize_handler() 
 
 Set objReminders = Application.Reminders 
 
End Sub 
 
 
 
Private Sub objReminders_ReminderRemove() 
 
'Occurs when a reminder is removed from the collection 
 
'or the user clicks Dismiss 
 
 
 
 MsgBox "A reminder has been removed from the collection." 
 
 
 
End Sub
```


## See also


#### Concepts


[Reminders Object](reminders-object-outlook.md)

