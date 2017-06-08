---
title: SelectNamesDialog.ShowOnlyInitialAddressList Property (Outlook)
keywords: vbaol11.chm833
f1_keywords:
- vbaol11.chm833
ms.prod: outlook
api_name:
- Outlook.SelectNamesDialog.ShowOnlyInitialAddressList
ms.assetid: 4159aa09-e790-523a-fd27-262d477599e3
ms.date: 06/08/2017
---


# SelectNamesDialog.ShowOnlyInitialAddressList Property (Outlook)

Returns or sets a  **Boolean** that determines if the **[AddressList](addresslist-object-outlook.md)** represented by **[SelectNamesDialog.InitialAddressList](selectnamesdialog-initialaddresslist-property-outlook.md)** is the only **AddressList** available in the drop-down list for **Address Book** in the **Select Names** dialog box. Read/write.


## Syntax

 _expression_ . **ShowOnlyInitialAddressList**

 _expression_ A variable that represents a **SelectNamesDialog** object.


## Remarks

The default value of this property is  **False** , meaning that all address lists are displayed. To restrict the drop-down list for **Address Book** to the one indicated by **InitialAddressList** , set **ShowOnlyInitialAddressList** to **True** .

If you do not set the  **InitialAddressList** property and then set **ShowOnlyInitialAddressList** to **True** , then the **AddressList** with **[AddressList.IsInitialAddressList](addresslist-isinitialaddresslist-property-outlook.md)** equal to **True** will be the only address list available in the drop-down list for **Address Book**.


## Example

The following code sample shows how to use  **IsInitialAddressList** and **ShowOnlyInitialAddressList** to have the **Select Names** dialog box always display only the address list in the default Contacts folder, regardless of the user's setting for the initial address list.


```vb
Sub ShowOnlyContacts() 
 
 Dim oMsg As MailItem 
 
 Set oMsg = Application.CreateItem(olMailItem) 
 
 
 
 Dim oDialog As SelectNamesDialog 
 
 Set oDialog = Application.Session.GetSelectNamesDialog 
 
 
 
 Dim oContacts As Folder 
 
 Set oContacts = _ 
 
 Application.Session.GetDefaultFolder(olFolderContacts) 
 
 
 
 Dim oAL As AddressList 
 
 For Each oAL In Application.Session.AddressLists 
 
 If oAL.GetContactsFolder = oContacts Then 
 
 Exit For 
 
 End If 
 
 Next 
 
 With oDialog 
 
 .InitialAddressList = oAL 
 
 .ShowOnlyInitialAddressList = True 
 
 .Recipients = oMsg.Recipients 
 
 If .Display Then 
 
 'Recipients Resolved 
 
 End If 
 
 End With 
 
End Sub
```


## See also


#### Concepts


[SelectNamesDialog Object](selectnamesdialog-object-outlook.md)

