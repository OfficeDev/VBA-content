---
title: Reminder.IsVisible Property (Outlook)
keywords: vbaol11.chm561
f1_keywords:
- vbaol11.chm561
ms.prod: outlook
api_name:
- Outlook.Reminder.IsVisible
ms.assetid: e99f8fab-32fa-94ef-be9b-523b580fa551
ms.date: 06/08/2017
---


# Reminder.IsVisible Property (Outlook)

Returns a  **Boolean** that determines if the reminder is currently visible. Read-only.


## Syntax

 _expression_ . **IsVisible**

 _expression_ A variable that represents a **Reminder** object.


## Remarks

 Outlook determines the return value of this property based on the state of the current reminder. All active reminders are visible. If **IsVisible** is **True** , the reminder is visible.


## Example

The following Microsoft Visual Basic for Applications (VBA) example dismisses all reminders that are currently visible. For example, if the current reminder is active, the  **IsVisible** property will return **True** .


```vb
Sub DismissReminders() 
 
 'Dismisses any active reminders. 
 
 Dim objRems As Outlook.Reminders 
 
 Dim objRem As Outlook.Reminder 
 
 Dim i As Integer 
 
 
 
 Set objRems = Application.Reminders 
 
 For i = objRems.Count To 1 Step -1 
 
 If objRems(i).IsVisible = True Then 
 
 objRems(i).Dismiss 
 
 End If 
 
 Next 
 
 Set olApp = Nothing 
 
 Set objRems = Nothing 
 
 Set objRem = Nothing 
 
End Sub
```


## See also


#### Concepts


[Reminder Object](reminder-object-outlook.md)

