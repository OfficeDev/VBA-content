---
title: Application.Reminders Property (Outlook)
keywords: vbaol11.chm731
f1_keywords:
- vbaol11.chm731
ms.prod: outlook
api_name:
- Outlook.Application.Reminders
ms.assetid: 1f5428f0-6362-a691-2fad-c80e48dce3f5
ms.date: 06/08/2017
---


# Application.Reminders Property (Outlook)

Returns a  **[Reminders](reminders-object-outlook.md)** collection that represents all current reminders. Read-only.


## Syntax

 _expression_ . **Reminders**

 _expression_ A variable that represents an **Application** object.


## Example

The following example returns the  **Reminders** collection and displays the captions of all reminders in the collection. If no current reminders are available, a message is displayed to the user.


```vb
Sub ViewReminderInfo() 
 
 'Lists reminder caption information 
 
 Dim objRem As Outlook.Reminder 
 
 Dim objRems As Outlook.Reminders 
 
 Dim strTitle As String 
 
 Dim strReport As String 
 
 
 
 Set objRems = Application.Reminders 
 
 strTitle = "Current Reminders:" 
 
 strReport = "" 
 
 'If there are reminders, display message 
 
 If Application.Reminders.Count <> 0 Then 
 
 For Each objRem In objRems 
 
 'Add information to string 
 
 strReport = strReport &; objRem.Caption &; vbCr 
 
 Next objRem 
 
 'Display report in dialog 
 
 MsgBox strTitle &; vbCr &; vbCr &; strReport 
 
 Else 
 
 MsgBox "There are no reminders in the collection." 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-outlook.md)

