---
title: MailItem.ReceivedByEntryID Property (Outlook)
keywords: vbaol11.chm1341
f1_keywords:
- vbaol11.chm1341
ms.prod: outlook
api_name:
- Outlook.MailItem.ReceivedByEntryID
ms.assetid: db4325d3-4442-220d-a812-1d3e4a0085bf
ms.date: 06/08/2017
---


# MailItem.ReceivedByEntryID Property (Outlook)

Returns a  **String** representing the **[EntryID](recipient-entryid-property-outlook.md)** for the true recipient as set by the transport provider delivering the mail message. Read-only.


## Syntax

 _expression_ . **ReceivedByEntryID**

 _expression_ A variable that represents a **MailItem** object.


## Remarks

This property corresponds to the MAPI property  **PidTagReceivedByEntryId** .

If you are getting this property in a Microsoft Visual Basic or Microsoft Visual Basic for Applications (VBA) solution, owing to some type issues, instead of directly referencing  **ReceivedByEntryID** , you should get the property through the **[PropertyAccessor](propertyaccessor-object-outlook.md)** object returned by the **[MailItem.PropertyAccessor](mailitem-propertyaccessor-property-outlook.md)** property, specifying the **PidTagReceivedByEntryId** property and its MAPI proptag namespace. The following code sample in VBA shows the workaround.




```vb
Public Sub GetReceiverEntryID() 
 
 Dim objInbox As Outlook.Folder 
 
 Dim objMail As Outlook.MailItem 
 
 Dim oPA As Outlook.PropertyAccessor 
 
 Dim strEntryID As String 
 
 Const PidTagReceivedByEntryId As String = "http://schemas.microsoft.com/mapi/proptag/0x003F0102" 
 
 
 
 Set objInbox = Application.Session.GetDefaultFolder(olFolderInbox) 
 
 Set objMail = objInbox.Items(1) 
 
 Set oPA = objMail.PropertyAccessor 
 
 strEntryID = oPA.BinaryToString(oPA.GetProperty(PidTagReceivedByEntryId)) 
 
 Debug.Print strEntryID 
 
 
 
 Set objInbox = Nothing 
 
 Set objMail = Nothing 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

