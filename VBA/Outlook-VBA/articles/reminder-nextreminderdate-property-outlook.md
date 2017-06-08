---
title: Reminder.NextReminderDate Property (Outlook)
keywords: vbaol11.chm563
f1_keywords:
- vbaol11.chm563
ms.prod: outlook
api_name:
- Outlook.Reminder.NextReminderDate
ms.assetid: c88a2606-fe30-d8c1-b16f-fd07b5596895
ms.date: 06/08/2017
---


# Reminder.NextReminderDate Property (Outlook)

Returns a  **Date** that indicates the next time the specified reminder will occur. Read-only.


## Syntax

 _expression_ . **NextReminderDate**

 _expression_ A variable that represents a **Reminder** object.


## Remarks

The  **NextReminderDate** property value changes every time the object's **[Snooze](reminder-snooze-method-outlook.md)** method is executed or when the user clicks the **Snooze** button.


## Example

The following example creates a report of all reminders in the collection and the dates when they will next occur. The subroutine concatenates the  **[Caption](reminder-caption-property-outlook.md)** and **NextReminderDate** properties into a string and displays the string in a dialog box.


```vb
Sub DisplayNextDateReport() 
 
 'Displays the next time all reminders will be displayed. 
 
 Dim objRems As Outlook.Reminders 
 
 Dim objRem As Outlook.Reminder 
 
 Dim strTitle As String 
 
 Dim strReport As String 
 
 
 
 Set objRems = Application.Reminders 
 
 strTitle = "Current Reminder Schedule:" 
 
 strReport = "" 
 
 'Check if any reminders exist. 
 
 If objRems.Count = 0 Then 
 
 MsgBox "There are no current reminders." 
 
 Else 
 
 For Each objRem In objRems 
 
 'Add information to string. 
 
 strReport = strReport &; objRem.Caption &; vbTab &; _ 
 
 objRem.NextReminderDate &; vbCr 
 
 Next objRem 
 
 'Display report in dialog box 
 
 MsgBox strTitle &; vbCr &; vbCr &; strReport 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Reminder Object](reminder-object-outlook.md)

