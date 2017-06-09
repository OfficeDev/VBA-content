---
title: Reminder.Dismiss Method (Outlook)
keywords: vbaol11.chm558
f1_keywords:
- vbaol11.chm558
ms.prod: outlook
api_name:
- Outlook.Reminder.Dismiss
ms.assetid: cc757453-5eab-4e9f-5dd2-2b7620506d11
ms.date: 06/08/2017
---


# Reminder.Dismiss Method (Outlook)

Dismisses the current reminder.


## Syntax

 _expression_ . **Dismiss**

 _expression_ A variable that represents a **Reminder** object.


## Remarks

The  **Dismiss** method will fail if there is no visible reminder.


## Example

The following example dismisses all active reminders. A reminder is active if its  **[IsVisible](reminder-isvisible-property-outlook.md)** property is set to **True** .


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

