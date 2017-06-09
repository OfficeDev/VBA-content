---
title: Display a Dialog Box for Selecting Entries from the Contacts Folder
ms.prod: outlook
ms.assetid: 6d31ad3e-8930-d571-3bfd-349efbf69232
ms.date: 06/08/2017
---


# Display a Dialog Box for Selecting Entries from the Contacts Folder

This topic describes how to use the  **[SelectNamesDialog](selectnamesdialog-object-outlook.md)** object to display entries from the Contacts folder in a dialog box that resembles the **Select Names** dialog box in the Outlook user interface.



1. Look for the address list that corresponds with the Contacts folder.The  **SelectNamesDialog** object displays entires in a dialog box based on an **[AddressList](addresslist-object-outlook.md)**. To display entries in the Contacts folder, look for the  **AddressList** that corresponds with the Contacts folder. Iterate through all the address lists defined for the current session, and for each address list, use **[AddressList.GetContactsFolder](addresslist-getcontactsfolder-method-outlook.md)**to match the corresponding folder with the Contacts folder. 
    
2. Initialize the dialog box with the address list of the Contacts folder.
    
3. Use  **[SelectNamesDialog.Display](selectnamesdialog-display-method-outlook.md)** to display the dialog box. If **SelectNamesDialog.Display** returns True, then selected entries will be available in **[SelectNamesDialog.Recipients](selectnamesdialog-recipients-property-outlook.md)**.
    




```vb
Sub ShowContactsInDialog() 
 Dim oDialog As SelectNamesDialog 
 Dim oAL As AddressList 
 Dim oContacts As Folder 
 
 Set oDialog = Application.Session.GetSelectNamesDialog 
 Set oContacts = _ 
 Application.Session.GetDefaultFolder(olFolderContacts) 
 
 'Look for the address list that corresponds with the Contacts folder 
 For Each oAL In Application.Session.AddressLists 
 If oAL.GetContactsFolder = oContacts Then 
 Exit For 
 End If 
 Next 
 With oDialog 
 'Initialize the dialog box with the address list representing the Contacts folder 
 .InitialAddressList = oAL 
 .ShowOnlyInitialAddressList = True 
 If .Display Then 
 'Recipients Resolved 
 'Access Recipients using oDialog.Recipients 
 End If 
 End With 
End Sub
```


