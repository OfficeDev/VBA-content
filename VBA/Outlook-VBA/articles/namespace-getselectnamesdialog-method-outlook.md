---
title: NameSpace.GetSelectNamesDialog Method (Outlook)
keywords: vbaol11.chm781
f1_keywords:
- vbaol11.chm781
ms.prod: outlook
api_name:
- Outlook.NameSpace.GetSelectNamesDialog
ms.assetid: 883d90e0-b3cc-e76e-cbe6-cb271e9ccb37
ms.date: 06/08/2017
---


# NameSpace.GetSelectNamesDialog Method (Outlook)

Obtains a  **[SelectNamesDialog](selectnamesdialog-object-outlook.md)** object for the current session.


## Syntax

 _expression_ . **GetSelectNamesDialog**

 _expression_ A variable that represents a **[NameSpace](namespace-object-outlook.md)** object.


### Return Value

A  **SelectNamesDialog** object for the current session. The **SelectNamesDialog** object supports displaying the **Select Names** dialog box for the user to select entries from one or more address lists in the current session.


## Example

The following code sample shows how to instantiate an instance of  **SelectNamesDialog** for the current session, and use it to display entries from the Contacts folder in a dialog box that resembles the **Select Names** dialog box in the Outlook user interface.


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


## See also


#### Concepts


[NameSpace Object](namespace-object-outlook.md)

