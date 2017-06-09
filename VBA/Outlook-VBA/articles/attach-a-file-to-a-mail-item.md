---
title: Attach a File to a Mail Item
ms.prod: outlook
ms.assetid: 1d94629b-e713-92cb-32de-c8910612e861
ms.date: 06/08/2017
---


# Attach a File to a Mail Item

This topic shows a procedure that attaches a spreadsheet file to a mail item. The procedure,  `AddAttachment`, assumes that the specified spreadsheet, Q496.xlsx, exists in the specified path, D:\Documents.  `AddAttachment` creates a new mail message, attaches Q496.xlsx to the mail message, assigns the attachment a descriptive caption, and displays the mail message with this attachment.


```vb
Sub AddAttachment() 
 Dim myItem As Outlook.MailItem 
 Dim myAttachments As Outlook.Attachments 
 
 Set myItem = Application.CreateItem(olMailItem) 
 Set myAttachments = myItem.Attachments 
 myAttachments.Add "D:\Documents\Q496.xlsx", _ 
 olByValue, 1, "4th Quarter 1996 Results Chart" 
 myItem.Display 
End Sub
```


