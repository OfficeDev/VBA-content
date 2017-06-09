---
title: ContactItem.Email3EntryID Property (Outlook)
keywords: vbaol11.chm1002
f1_keywords:
- vbaol11.chm1002
ms.prod: outlook
api_name:
- Outlook.ContactItem.Email3EntryID
ms.assetid: f38c8002-c4a8-f47a-c783-986e4121f4c3
ms.date: 06/08/2017
---


# ContactItem.Email3EntryID Property (Outlook)

Returns a  **String** representing the entry ID of the third e-mail entry for the contact. Read-only.


## Syntax

 _expression_ . **Email3EntryID**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property corresponds to the MAPI named property  **dispidEmail3OriginalEntryID** .

If you are getting this property in a Microsoft Visual Basic or Microsoft Visual Basic for Applications (VBA) solution, owing to some type issues, instead of directly referencing  **Email3EntryID** , you should get the property through the **[PropertyAccessor](propertyaccessor-object-outlook.md)** object returned by the **[ContactItem.PropertyAccessor](contactitem-propertyaccessor-property-outlook.md)** property, specifying the MAPI property **PidLidEmail3OriginalEntryId** property and its MAPI id namespace. The following code sample in VBA shows the workaround.




```vb
Public Sub GetEmail3EntryID() 
 
 Dim objContactFolder As Outlook.Folder 
 
 Dim objContactItem As Outlook.ContactItem 
 
 Dim objRec As Outlook.Recipient 
 
 Dim strEntryID As String 
 
 Dim oPA As Outlook.PropertyAccessor 
 
 Const EMAIL3_ENTRYID As String = "http://schemas.microsoft.com/mapi/id/{00062004-0000-0000-C000-000000000046}/80A50102" 
 
 
 
 Set objContactFolder = Application.Session.GetDefaultFolder(olFolderContacts) 
 
 Set objContactItem = objContactFolder.Items(1) 
 
 Set oPA = objContactItem.PropertyAccessor 
 
 strEntryID = oPA.BinaryToString(oPA.GetProperty(EMAIL3_ENTRYID)) 
 
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

