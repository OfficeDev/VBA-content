---
title: AddressList.GetContactsFolder Method (Outlook)
keywords: vbaol11.chm2032
f1_keywords:
- vbaol11.chm2032
ms.prod: outlook
api_name:
- Outlook.AddressList.GetContactsFolder
ms.assetid: 9ea91624-bd7d-af64-7220-a2d9b659787a
ms.date: 06/08/2017
---


# AddressList.GetContactsFolder Method (Outlook)

Obtains a  **[Folder](folder-object-outlook.md)** object that represents the Contacts folder for the **[AddressList](addresslist-object-outlook.md)** object.


## Syntax

 _expression_ . **GetContactsFolder**

 _expression_ A variable that represents an **AddressList** object.


### Return Value

A  **Folder** object that represents the Outlook Contacts folder for the **AddressList** . Returns **Null** ( **Nothing** in Visual Basic) if no Outlook contacts folder is found.


## Remarks

This method allows you to match an  **AddressList** for the Contacts folder that you would like to set up as the initial address list in the **Select Names** dialog box.


## Example

The following code sample shows you how to initialize the  **Select Names** dialog box with the **AddressList** for the default Contacts folder. It first obtains the **Folder** object for the default Contacts folder, and looks for its **AddressList** by comparing the Entry ID of this **Folder** object with the Entry ID of the **Folder** object assoicated with each **AddressList** in the current session until it finds a match. It then sets the **[InitialAddressList](selectnamesdialog-initialaddresslist-property-outlook.md)** property and displays the **Select Names** dialog box.


```vb
Sub SetContactsFolderAsInitialAddressList() 
 
 Dim oMsg As MailItem 
 
 Set oMsg = Application.CreateItem(olMailItem) 
 
 Dim oDialog As SelectNamesDialog 
 
 Set oDialog = Application.Session.GetSelectNamesDialog 
 
 Dim oAL As AddressList 
 
 Dim oContacts As Folder 
 
 Set oContacts = _ 
 
 Application.Session.GetDefaultFolder(olFolderContacts) 
 
 
 
 On Error GoTo HandleError 
 
 'Look for the AddressList for the default Contacts folder 
 
 For Each oAL In Application.Session.AddressLists 
 
 If oAL.AddressListType = olOutlookAddressList Then 
 
 If oAL.GetContactsFolder.EntryID = _ 
 
 oContacts.EntryID Then 
 
 Exit For 
 
 End If 
 
 End If 
 
 Next 
 
 
 
 With oDialog 
 
 .Caption = "Select Customer Contact" 
 
 .ToLabel = "Customer C&;ontact" 
 
 .NumberOfRecipientSelectors = olShowTo 
 
 .InitialAddressList = oAL 
 
 
 
 'Let the selected names be the recipients of the new message 
 
 .Recipients = oMsg.Recipients 
 
 
 
 If .Display Then 
 
 'Recipients Resolved 
 
 End If 
 
 End With 
 
 
 
HandleError: 
 
 Exit Sub 
 
End Sub
```


## See also


#### Concepts


[AddressList Object](addresslist-object-outlook.md)

