---
title: SelectNamesDialog.InitialAddressList Property (Outlook)
keywords: vbaol11.chm835
f1_keywords:
- vbaol11.chm835
ms.prod: outlook
api_name:
- Outlook.SelectNamesDialog.InitialAddressList
ms.assetid: 19cfe6be-e6b5-62e0-741a-b196ef7bac77
ms.date: 06/08/2017
---


# SelectNamesDialog.InitialAddressList Property (Outlook)

Returns or sets an  **[AddressList](addresslist-object-outlook.md)** object that determines the initial address list to be displayed in the **Select Names** dialog box. Read/write.


## Syntax

 _expression_ . **InitialAddressList**

 _expression_ A variable that represents a **SelectNamesDialog** object.


## Remarks

Setting the  **InitialAddressList** property is the programmatic equivalent to selecting an **AddressList** from the drop-down list for **Address Book** in the **Select Names** dialog box.

In its default state,  **InitialAddressList** is the **AddressList** that has the property **[AddressList.IsInitialAddressList](addresslist-isinitialaddresslist-property-outlook.md)** set to **True** . **IsInitialAddressList** corresponds to setting **Show this address list first** in the **Addressing** dialog box, which is available by clicking **Tools**, and then  **Options** in the **Address Book** dialog box.


## Example

The following code sample shows how to use  **InitialAddressList** and **[SelectNamesDialog.ShowOnlyInitialAddressList](selectnamesdialog-showonlyinitialaddresslist-property-outlook.md)** to have the **Select Names** dialog box always display only the address list in the default Contacts folder, regardless of the user's setting for the initial address list.


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

