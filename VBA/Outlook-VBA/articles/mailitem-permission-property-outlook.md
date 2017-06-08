---
title: MailItem.Permission Property (Outlook)
keywords: vbaol11.chm1386
f1_keywords:
- vbaol11.chm1386
ms.prod: outlook
api_name:
- Outlook.MailItem.Permission
ms.assetid: 394173d4-344a-148a-1628-b4ca47d4ef2d
ms.date: 06/08/2017
---


# MailItem.Permission Property (Outlook)

Sets or returns an  **[OlPermission](olpermission-enumeration-outlook.md)** constant that determines what permissions to grant to the recipients of the e-mail item. Read/write.


## Syntax

 _expression_ . **Permission**

 _expression_ A variable that represents a **[MailItem](mailitem-object-outlook.md)** object.


## Remarks

The  **Permission** property should be synchronized with the **[PermissionTemplateGuid](mailitem-permissiontemplateguid-property-outlook.md)** property to accurately reflect the permission status of the **MailItem** . Setting the **PermissionTemplateGuid** property to a valid GUID also sets the **Permission** property to **OlPermission.olPermissionTemplate** .

 When no Information Rights Management (IRM) has been set up, (in which case the **Permission** property is **OlPermission.olUnrestricted** ), or the restriction is not to forward the **MailItem** , (in which case the **Permission** property is **OlPermission.olDoNotForward** ), the value of the **PermissionTemplateGuid** property should be an empty string.

Although you can view content that is protected by IRM on any computer that is running the 2007 Microsoft Office system or a later version, you must have Microsoft Office Professional Edition 2003, Microsoft Office Outlook 2007, or a later version of Outlook to create or send an e-mail that is protected by IRM.


## Example

This Microsoft Visual Basic for Applications (VBA) example uses the  **[Send](mailitem-send-method-outlook.md)** event and sends an item with a 'Do not forward' restriction. You must place the sample code in a class module such as **ThisOutlookSession**, and the  `SendMyMail` procedure must be called before the event procedure can be called by Microsoft Outlook. Replace 'Dan Wilson' with a valid recipient name before you run this example.


```vb
Public WithEvents myItem As Outlook.MailItem 
 
 
 
Sub SendMyMail() 
 
 Set myItem = Outlook.CreateItem(olMailItem) 
 
 myItem.To = "Dan Wilson" 
 
 myItem.Subject = "Data files information" 
 
 myItem.Send 
 
End Sub 
 
 
 
Private Sub myItem_Send(Cancel As Boolean) 
 
 myItem.Permission = olDoNotForward 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

