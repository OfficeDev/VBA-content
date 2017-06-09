---
title: Reminder.OriginalReminderDate Property (Outlook)
keywords: vbaol11.chm564
f1_keywords:
- vbaol11.chm564
ms.prod: outlook
api_name:
- Outlook.Reminder.OriginalReminderDate
ms.assetid: ecc3f0c4-0e20-1d02-94b5-40807523ad2d
ms.date: 06/08/2017
---


# Reminder.OriginalReminderDate Property (Outlook)

Returns a  **Date** that specifies the original date and time that the specified reminder is set to occur. Read-only.


## Syntax

 _expression_ . **OriginalReminderDate**

 _expression_ A variable that represents a **Reminder** object.


## Remarks

This value corresponds to the original date and time value before the  **[Snooze](reminder-snooze-method-outlook.md)** method is executed or the user clicks the **Snooze** button.


## Example

The following Microsoft Visual Basic for Applications (VBA) example creates a report of all reminders in the  **[Reminders](reminders-object-outlook.md)** collection and the dates at which they are scheduled to occur. The subroutine concatenates the **[Caption](reminder-caption-property-outlook.md)** and **OriginalReminderDate** properties of all **[Reminder](reminder-object-outlook.md)** objects in the collection into a string and displays the string in a dialog box.


```vb
Sub DisplayOriginalDateReport() 
 
 'Displays the time at which all reminders will be displayed. 
 
 Dim objRems As Outlook.Reminders 
 
 Dim objRem As Outlook.Reminder 
 
 Dim strTitle As String 
 
 Dim strReport As String 
 
 
 
 Set objRems = Application.Reminders 
 
 strTitle = "Original Reminder Schedule:" 
 
 strReport = "" 
 
 'Check if any reminders exist. 
 
 If objRems.Count = 0 Then 
 
 MsgBox "There are no current reminders." 
 
 Else 
 
 For Each objRem In objRems 
 
 'Add info to string 
 
 strReport = strReport &; objRem.Caption &; vbTab &; vbTab &; _ 
 
 objRem.OriginalReminderDate &; vbCr 
 
 Next objRem 
 
 'Display report in dialog 
 
 MsgBox strTitle &; vbCr &; vbCr &; strReport 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Reminder Object](reminder-object-outlook.md)

