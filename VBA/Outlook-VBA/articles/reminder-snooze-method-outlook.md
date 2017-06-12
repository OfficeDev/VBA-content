---
title: Reminder.Snooze Method (Outlook)
keywords: vbaol11.chm559
f1_keywords:
- vbaol11.chm559
ms.prod: outlook
api_name:
- Outlook.Reminder.Snooze
ms.assetid: bb417d32-d69b-7f9d-4ca3-b85888421e7b
ms.date: 06/08/2017
---


# Reminder.Snooze Method (Outlook)

Delays the reminder by a specified time. 


## Syntax

 _expression_ . **Snooze**( **_SnoozeTime_** )

 _expression_ An expression that returns a **Reminder** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SnoozeTime_|Optional| **Variant**|Indicates the amount of time (in minutes) to delay the reminder. The default value is 5 minutes.|

## Remarks

This is equivalent to the user clicking the  **Snooze** button.

This method will fail if the current reminder is not active.


## Example

The following Microsoft Visual Basic for Applications (VBA) example delays all active reminders by a specified amount of time.


```vb
Sub SnoozeReminders() 
 
 'Delays all reminders by a specified amount of time 
 
 Dim objRems As Outlook.Reminders 
 
 Dim objRem As Outlook.Reminder 
 
 Dim varTime As Variant 
 
 
 
 Set objRems = Application.Reminders 
 
 varTime = InputBox("Type the number of minutes to delay") 
 
 For Each objRem In objRems 
 
 If objRem.IsVisible = True Then 
 
 objRem.Snooze (varTime) 
 
 End If 
 
 Next objRem 
 
End Sub
```


## See also


#### Concepts


[Reminder Object](reminder-object-outlook.md)

