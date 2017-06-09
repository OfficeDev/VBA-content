---
title: MailItem.PermissionService Property (Outlook)
keywords: vbaol11.chm1387
f1_keywords:
- vbaol11.chm1387
ms.prod: outlook
api_name:
- Outlook.MailItem.PermissionService
ms.assetid: c999b215-f360-17b1-4915-45c3b525d3e5
ms.date: 06/08/2017
---


# MailItem.PermissionService Property (Outlook)

Sets or returns an  **[OlPermissionService](olpermissionservice-enumeration-outlook.md)** constant that determines the permission service that will be used when sending a message protected by Information Rights Management (IRM). Read/write.


## Syntax

 _expression_ . **PermissionService**

 _expression_ A variable that represents a **MailItem** object.


## Remarks

This property is useful only if you have more than one permission identity for a particular SMTP address. 

While you can view content that is protected by IRM on any computer running the 2007 Microsoft Office system or a later version, you must have Microsoft Office Professional Edition 2003, Microsoft Office Outlook 2007, or a later version of Outlook to create or send an e-mail that is protected by IRM.


## Example

This Microsoft Visual Basic for Applications (VBA) example demonstrates how to specify the permission service before sending an item. Replace 'Dan Wilson' with a valid recipient name before running this example.


```vb
Sub SendMyMail() 
 
 Set myItem = Outlook.CreateItem(olMailItem) 
 
 myItem.To = "Dan Wilson" 
 
 myItem.Subject = "Data files information" 
 
 myItem.Permission = olDoNotForward 
 
 myItem.PermissionService = olWindows 
 
 myItem.Send 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

