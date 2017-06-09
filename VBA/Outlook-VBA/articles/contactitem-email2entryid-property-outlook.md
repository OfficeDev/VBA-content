---
title: ContactItem.Email2EntryID Property (Outlook)
keywords: vbaol11.chm998
f1_keywords:
- vbaol11.chm998
ms.prod: outlook
api_name:
- Outlook.ContactItem.Email2EntryID
ms.assetid: 0c5691bb-e112-763b-d126-2bcc2c52ccce
ms.date: 06/08/2017
---


# ContactItem.Email2EntryID Property (Outlook)

Returns a  **String** representing the entry ID of the second e-mail entry for the contact. Read-only.


## Syntax

 _expression_ . **Email2EntryID**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property corresponds to the MAPI named property  **dispidEmail2OriginalEntryID** .

If you are getting this property in a Microsoft Visual Basic or Microsoft Visual Basic for Applications (VBA) solution, owing to some type issues, instead of directly referencing  **Email2EntryID** , you should get the property through the **[PropertyAccessor](propertyaccessor-object-outlook.md)** object returned by the **[ContactItem.PropertyAccessor](contactitem-propertyaccessor-property-outlook.md)** property, specifying the MAPI property **PidLidEmail2OriginalEntryId** property and its MAPI id namespace. The following code sample in VBA shows the workaround.




```vb
Public Sub GetEmail2EntryID() 
 
 Dim objContactFolder As Outlook.Folder 
 
 Dim objContactItem As Outlook.ContactItem 
 
 Dim objRec As Outlook.Recipient 
 
 Dim strEntryID As String 
 
 Dim oPA As Outlook.PropertyAccessor 
 
 Const EMAIL2_ENTRYID As String = "http://schemas.microsoft.com/mapi/id/{00062004-0000-0000-C000-000000000046}/80950102" 
 
 
 
 Set objContactFolder = Application.Session.GetDefaultFolder(olFolderContacts) 
 
 Set objContactItem = objContactFolder.Items(1) 
 
 Set oPA = objContactItem.PropertyAccessor 
 
 strEntryID = oPA.BinaryToString(oPA.GetProperty(EMAIL2_ENTRYID)) 
 
 Debug.Print strEntryID 
 
 Set objRec = Application.Session.GetRecipientFromID(strEntryID) 
 
 If objRec Is Nothing Then 
 
 Debug.Print "GetRecipientFromID failed" 
 
 Else 
 
 Debug.Print objRec.Name 
 
 Debug.Print objRec.EntryID 
 
 End If 
 
 
 
 'Cleanup 
 
 Set objContactItem = Nothing 
 
 Set objContactFolder = Nothing 
 
End Sub
```


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

