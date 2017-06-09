---
title: Reminders Object (Outlook)
keywords: vbaol11.chm3015
f1_keywords:
- vbaol11.chm3015
ms.prod: outlook
api_name:
- Outlook.Reminders
ms.assetid: 66b94251-7fe4-886b-7c29-7feac4440dee
ms.date: 06/08/2017
---


# Reminders Object (Outlook)

Contains a collection of all the  **[Reminder](reminder-object-outlook.md)** objects in a Microsoft Outlook application that represent the reminders for all pending items.


## Remarks

Use the  **[Application](application-object-outlook.md)** object's **[Reminders](application-reminders-property-outlook.md)** property to return the **Reminders** collection. Use **Reminders** ( _index_ ), where _index_ is the name or ordinal value of the reminder, to return a single **[Reminder](reminder-object-outlook.md)** object.

Reminders are created programmatically when a new Outlook item is created with a reminder. For example, a reminder is created when an  **[AppointmentItem](appointmentitem-object-outlook.md)** object is created and the **AppointmentItem** object's **[ReminderSet](appointmentitem-reminderset-property-outlook.md)** property is set to **True**.


## Example

The following example displays the captions of each reminder in the list.


```
Sub ViewReminderInfo() 
 'Lists reminder caption information 
 Dim objRem As Reminder 
 Dim objRems As Reminders 
 Dim strTitle As String 
 Dim strReport As String 
 
 Set objRems = Application.Reminders 
 strTitle = "Current Reminders:" 
 'If there are reminders, display message 
 If Application.Reminders.Count <> 0 Then 
 For Each objRem In objRems 
 'If string is empty, create new string 
 If strReport = "" Then 
 strReport = objRem.Caption &amp; vbCr 
 Else 
 'Add info to string 
 strReport = strReport &amp; objRem.Caption &amp; vbCr 
 End If 
 Next objRem 
 'Display report in dialog 
 MsgBox strTitle &amp; vbCr &amp; vbCr &amp; strReport 
 Else 
 MsgBox "There are no reminders in the collection." 
 End If 
End Sub
```

The following example creates a new appointment item and sets the  **ReminderSet** property to **True**, adding a new **Reminder** object to the **Reminders** collection.




```
Sub AddAppt() 
 'Adds a new appointment and reminder to the reminders collection 
 Dim objApt As AppointmentItem 
 
 Set objApt = Application.CreateItem(olAppointmentItem) 
 objApt.ReminderSet = True 
 objApt.Subject = "Tuesday's meeting" 
 objApt.Save 
End Sub
```


## Events



|**Name**|
|:-----|
|[BeforeReminderShow](reminders-beforeremindershow-event-outlook.md)|
|[ReminderAdd](reminders-reminderadd-event-outlook.md)|
|[ReminderChange](reminders-reminderchange-event-outlook.md)|
|[ReminderFire](reminders-reminderfire-event-outlook.md)|
|[ReminderRemove](reminders-reminderremove-event-outlook.md)|
|[Snooze](reminders-snooze-event-outlook.md)|

## Methods



|**Name**|
|:-----|
|[Item](reminders-item-method-outlook.md)|
|[Remove](reminders-remove-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](reminders-application-property-outlook.md)|
|[Class](reminders-class-property-outlook.md)|
|[Count](reminders-count-property-outlook.md)|
|[Parent](reminders-parent-property-outlook.md)|
|[Session](reminders-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
