---
title: SelectNamesDialog Object (Outlook)
keywords: vbaol11.chm3156
f1_keywords:
- vbaol11.chm3156
ms.prod: outlook
api_name:
- Outlook.SelectNamesDialog
ms.assetid: 1522736a-3cad-9f1c-4da9-b52a3a01731c
ms.date: 06/08/2017
---


# SelectNamesDialog Object (Outlook)

Displays the  **Select Names** dialog box for the user to select entries from one or more address lists, and returns the selected entries in the collection object specified by the property **[SelectNamesDialog.Recipients](selectnamesdialog-recipients-property-outlook.md)**.


## Remarks

You can instantiate an instance of the  **SelectNamesDialog** object by calling **[NameSpace.GetSelectNamesDialog](namespace-getselectnamesdialog-method-outlook.md)**.

The dialog box displayed by  **[SelectNamesDialog.Display](selectnamesdialog-display-method-outlook.md)** is similar to the **Select Names** dialog box in the Outlook user interface. It observes the size and position settings of the built-in **Select Names** dialog box. However, its default state does not show **Message Recipients** above the **To**,  **Cc**, and  **Bcc** edit boxes. For more information on using the **SelectNamesDialog** object to display the **Select Names** dialog box, see[Display Names from the Address Book](http://msdn.microsoft.com/library/32e7179c-8133-ee20-ecf6-52c9275f205f%28Office.15%29.aspx).


## Example

The following code sample shows how to use the  **SelectNamesDialog** object to display entries from the Contacts folder in a dialog box that resembles the **Select Names** dialog box in the Outlook user interface.


```
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


## Methods



|**Name**|
|:-----|
|[Display](selectnamesdialog-display-method-outlook.md)|
|[SetDefaultDisplayMode](selectnamesdialog-setdefaultdisplaymode-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[AllowMultipleSelection](selectnamesdialog-allowmultipleselection-property-outlook.md)|
|[Application](selectnamesdialog-application-property-outlook.md)|
|[BccLabel](selectnamesdialog-bcclabel-property-outlook.md)|
|[Caption](selectnamesdialog-caption-property-outlook.md)|
|[CcLabel](selectnamesdialog-cclabel-property-outlook.md)|
|[Class](selectnamesdialog-class-property-outlook.md)|
|[ForceResolution](selectnamesdialog-forceresolution-property-outlook.md)|
|[InitialAddressList](selectnamesdialog-initialaddresslist-property-outlook.md)|
|[NumberOfRecipientSelectors](selectnamesdialog-numberofrecipientselectors-property-outlook.md)|
|[Parent](selectnamesdialog-parent-property-outlook.md)|
|[Recipients](selectnamesdialog-recipients-property-outlook.md)|
|[Session](selectnamesdialog-session-property-outlook.md)|
|[ShowOnlyInitialAddressList](selectnamesdialog-showonlyinitialaddresslist-property-outlook.md)|
|[ToLabel](selectnamesdialog-tolabel-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
