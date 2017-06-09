---
title: Reminder Object (Outlook)
keywords: vbaol11.chm3014
f1_keywords:
- vbaol11.chm3014
ms.prod: outlook
api_name:
- Outlook.Reminder
ms.assetid: b7364e48-51bc-b360-2154-e85e7779ece4
ms.date: 06/08/2017
---


# Reminder Object (Outlook)

Represents an Outlook reminder.


## Remarks

Reminders allow users to keep track of upcoming appointments by scheduling a pop-up dialog box to appear at a given time. In addition to appointments, reminders can occur for tasks, contacts and e-mail messages.

Use  **[Reminders](http://msdn.microsoft.com/library/1f5428f0-6362-a691-2fad-c80e48dce3f5%28Office.15%29.aspx)** ( _index_ ), where _index_ is the name or index number of the reminder, to return a single **Reminder** object.

Reminders are created programmatically when a new Microsoft Outlook item, such as an  **[AppointmentItem](appointmentitem-object-outlook.md)** object, is created and the item 's **[ReminderSet](http://msdn.microsoft.com/library/575d5fb2-1672-ddae-832c-7dcc7d1da2d6%28Office.15%29.aspx)** property is set to **True**.

Use the  **Reminders** collection's **[Remove](http://msdn.microsoft.com/library/c7a25177-8869-39c2-4109-5c2e2a4bd193%28Office.15%29.aspx)** method to remove a **Reminder** object from the collection. Once a reminder is removed from its associated item, the **AppointmentItem** object's **ReminderSet** property is set to **False**.


## Example

The following example displays the caption of the first reminder in the collection.


```
Sub ViewReminderInfo() 
 
 'Displays information about first reminder in collection 
 
 
 
 Dim colReminders As Outlook.Reminders 
 
 Dim objRem As Reminder 
 
 
 
 Set colReminders = Application.Reminders 
 
 'If there are reminders, display message 
 
 If colReminders.Count <> 0 Then 
 
 Set objRem = colReminders.Item(1) 
 
 MsgBox "The caption of the first reminder in the collection is: " &amp; _ 
 
 objRem.Caption 
 
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


## Methods



|**Name**|
|:-----|
|[Dismiss](http://msdn.microsoft.com/library/cc757453-5eab-4e9f-5dd2-2b7620506d11%28Office.15%29.aspx)|
|[Snooze](http://msdn.microsoft.com/library/bb417d32-d69b-7f9d-4ca3-b85888421e7b%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/cfbb4642-250b-96b9-492a-508f8309b867%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/b83b10f7-745c-337c-182b-74dabac65a17%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/b6178afe-19e9-5298-5624-f9c383ff4dd3%28Office.15%29.aspx)|
|[IsVisible](http://msdn.microsoft.com/library/e99f8fab-32fa-94ef-be9b-523b580fa551%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/f8fb20c5-bb36-73c0-d7c3-252307e96140%28Office.15%29.aspx)|
|[NextReminderDate](http://msdn.microsoft.com/library/c88a2606-fe30-d8c1-b16f-fd07b5596895%28Office.15%29.aspx)|
|[OriginalReminderDate](http://msdn.microsoft.com/library/ecc3f0c4-0e20-1d02-94b5-40807523ad2d%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/fdaa18ca-02ee-a5c4-ee8f-79da8db7447e%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/30bd8c36-1afa-aae1-f050-47ad43af53f9%28Office.15%29.aspx)|

## See also


#### Other resources


[Reminder Object Members](http://msdn.microsoft.com/library/2dc26aef-9636-4761-4d79-4571bb7c9726%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
