---
title: MailItem.ReceivedOnBehalfOfEntryID Property (Outlook)
keywords: vbaol11.chm1343
f1_keywords:
- vbaol11.chm1343
ms.prod: outlook
api_name:
- Outlook.MailItem.ReceivedOnBehalfOfEntryID
ms.assetid: fffcb637-9a7d-3541-49fc-85f314cd92cb
ms.date: 06/08/2017
---


# MailItem.ReceivedOnBehalfOfEntryID Property (Outlook)

Returns a  **String** representing the **[EntryID](recipient-entryid-property-outlook.md)** of the user delegated to represent the recipient for the mail message. Read-only.


## Syntax

 _expression_ . **ReceivedOnBehalfOfEntryID**

 _expression_ A variable that represents a **MailItem** object.


## Remarks

This property corresponds to the MAPI property  **PidTagReceivedRepresentingEntryId** .

If you are getting this property in a Microsoft Visual Basic or Microsoft Visual Basic for Applications (VBA) solution, owing to some type issues, instead of directly referencing  **ReceivedOnBehalfOfEntryID** , you should get the property through the **[PropertyAccessor](propertyaccessor-object-outlook.md)** object returned by the **[MailItem.PropertyAccessor](mailitem-propertyaccessor-property-outlook.md)** property, specifying the MAPI property **PidTagReceivedRepresentingEntryId** property and its MAPI proptag namespace. The following code sample in VBA shows the workaround.




```vb
Public Sub GetReceiverEntryID() 
 
 Dim objInbox As Outlook.Folder 
 
 Dim objMail As Outlook.MailItem 
 
 Dim oPA As Outlook.PropertyAccessor 
 
 Dim strEntryID As String 
 
 Const PidTagReceivedRepresentingEntryId As String = "http://schemas.microsoft.com/mapi/proptag/0x00430102" 
 
 
 
 Set objInbox = Application.Session.GetDefaultFolder(olFolderInbox) 
 
 Set objMail = objInbox.Items(1) 
 
 Set oPA = objMail.PropertyAccessor 
 
 strEntryID = oPA.BinaryToString(oPA.GetProperty(PidTagReceivedRepresentingEntryId)) 
 
 Debug.Print strEntryID 
 
 
 
 Set objInbox = Nothing 
 
 Set objMail = Nothing 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

