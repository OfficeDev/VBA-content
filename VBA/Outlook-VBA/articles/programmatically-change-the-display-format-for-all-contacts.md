---
title: Programmatically Change the Display Format for All Contacts
ms.prod: outlook
ms.assetid: 3cf2408e-4d9d-4b41-5cd0-1f3c12784fd4
ms.date: 06/08/2017
---


# Programmatically Change the Display Format for All Contacts

In Outlook, you can specify the default setting for how to file new contacts through the options for Contacts, as well as when you create the new contact. For example, the default setting is  **Last, First**, which files a contact by the last name followed by the first name. However, changing this setting only applies to new contacts that you create. For contacts that already exist, if you want to change the way that their names are filed by, for example, changing from the default  **Last, First** to **First, Last**, you will have to either do it individually for each existing contact in the inspector, or, you will have to write a macro to change the setting for all existing contacts in the Contacts folder.

This topic shows a code sample that goes through all the Contact items in the default Contacts folder and uses the  **[FileAs](contactitem-fileas-property-outlook.md)** property of each Contact item to specify the string to file the contact by; in this particular example, the string is changed to first name followed by a blank and then the last name. The code sample then saves the changes to the Contact item.

 **Note**  Generally, a folder in Outlook can contain heterogeneous items, and the Contact folder can contain  **[ContactItem](contactitem-object-outlook.md)** objects as well as other items. The code sample ensures that it only changes the file-as format for Contact items by filtering on the message class IPM.Contact. For more information on item types and message classes, see [Item Types and Message Classes](item-types-and-message-classes.md).




```vb
Private Sub ReFileContacts() 
 Dim items As items, item As ContactItem, folder As folder 
 Dim contactItems As Outlook.items 
 Dim itemContact As Outlook.ContactItem 
 
 Set folder = Session.GetDefaultFolder(olFolderContacts) 
 Set items = folder.items 
 Count = items.Count 
 If Count = 0 Then 
 MsgBox "Nothing to do!" 
 Exit Sub 
 End If 
 
 'Filter on the message class to obtain only contact items in the folder 
 Set contactItems = items.Restrict("[MessageClass]='IPM.Contact'") 
 
 For Each itemContact In contactItems 
 itemContact.FileAs = itemContact.FirstName + " " + itemContact.LastName 
 itemContact.Save 
 Next 
 
 MsgBox "Your contacts have been re-filed." 
End Sub
```


