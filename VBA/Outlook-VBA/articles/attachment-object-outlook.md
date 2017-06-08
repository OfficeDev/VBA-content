---
title: Attachment Object (Outlook)
keywords: vbaol11.chm2360
f1_keywords:
- vbaol11.chm2360
ms.prod: outlook
api_name:
- Outlook.Attachment
ms.assetid: 3e11582b-ac90-0948-bc37-506570bb287b
ms.date: 06/08/2017
---


# Attachment Object (Outlook)

Represents a document or link to a document contained in an Outlook item.


## Remarks

Use  **[Attachments](attachments-item-method-outlook.md)** ( _index_ ), where _index_ is the index number, to return a single **Attachment** object.

Use the  **[Add](attachments-add-method-outlook.md)** method to add an attachment to an item.


## Example

The following Visual Basic for Applications (VBA) example creates a new mail message, attaches Q496.xlsx as an attachment (not a link), assigns the attachment a descriptive caption, and displays the mail message with this attachment. This example assumes that the specified spreadsheet, Q496.xlsx, exists in the specified path, D:\Documents.


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


## See also


#### Other resources


[Attach a File to a Mail Item](http://msdn.microsoft.com/library/1d94629b-e713-92cb-32de-c8910612e861%28Office.15%29.aspx)
[Attach an Outlook Contact Item to an Email Message](http://msdn.microsoft.com/library/ae5240ad-dc3e-4499-8fd0-d8c2d90aa9ba%28Office.15%29.aspx)
[Limit the Size of an Attachment to an Outlook Email Message](http://msdn.microsoft.com/library/9a240e17-f715-482c-9a8b-c6be1144e15a%28Office.15%29.aspx)
[Modify an Attachment of an Outlook Email Message](http://msdn.microsoft.com/library/f5dac09a-272b-49d6-bf1e-82c3981260ed%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)


