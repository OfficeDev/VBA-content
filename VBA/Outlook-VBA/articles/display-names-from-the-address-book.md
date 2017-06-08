---
title: Display Names from the Address Book
ms.prod: outlook
ms.assetid: 32e7179c-8133-ee20-ecf6-52c9275f205f
ms.date: 06/08/2017
---


# Display Names from the Address Book

This topic describes the address book and explains how to programmatically display names from an address book in the Outlook Address Book dialog box.

Outlook maintains a hierarchy of address books for a session. In order for the transport provider to deliver a message, the recipient must have an address entry in one of the address books in this hierarchy. 

An address book contains one or more address lists. Each address list is composed of users, distribution lists, or other types of address entries. An example of an address list is the Exchange Global Address List. In the Outlook user interface, you can open the Address Book dialog box to view and select names in an address list. When you create a mail item or appointment item, or when you assign a task item, you can use the Address Book to help you select recipients. 

The Outlook Address Book is an address list or a set of address lists that Outlook creates automatically. By default, it contains one address list for the contacts in your Contacts folder that have at least one e-mail address or fax number entry. As you create other folders in the Contacts module, by default, each folder will also become an address list in the Outlook Address Book. In the  **Outlook Address Book** tab of the **Properties** dialog box for the folder, you can select or de-select **Show this folder as an e-mail address book**.

From the programmability perspective, Outlook maintains a collection of  **[AddressLists](addresslists-object-outlook.md)** for the current session. Each **[AddressList](addresslist-object-outlook.md)** consists of a collection of **[AddressEntries](addressentries-object-outlook.md)**. There are different types of address lists (as enumerated by  **[OlAddressListType](oladdresslisttype-enumeration-outlook.md)**) and different types of address entries (as enumerated by  **[OlAddressEntryUserType](oladdressentryusertype-enumeration-outlook.md)**). When you add a folder to the Contacts module, you can use  **[Folder.ShowAsOutlookAB](folder-showasoutlookab-property-outlook.md)** to specify whether that folder will be displayed as an address list in the Outlook Address Book.

The  **[Recipient](recipient-object-outlook.md)** object is associated with an **[AddressEntry](addressentry-object-outlook.md)** object that is specified by the **[Recipient.Address](recipient-address-property-outlook.md)** property. You can also use the **[AddressEntry.AddressEntryUserType](addressentry-addressentryusertype-property-outlook.md)** property to identify the type of the recipient, for example, whether the recipient is a Contact item, an Exchange user, or an Exchange distribution list.

The  **[SelectNamesDialog](selectnamesdialog-object-outlook.md)** object allows you to display names from an address list in a dialog box that resembles the **Select Names** dialog box in the Outlook user interface. The following figure is an example of the **Select Names** dialog box displaying the Contacts folder.

The dialog box allows a user to select entries from one or more address lists in the Address Book, and returns the selected recipients in the  **[SelectNamesDialog.Recipients](selectnamesdialog-recipients-property-outlook.md)** property. Through properties and methods of **SelectNamesDialog**, you can control the following aspects of the dialog box:


- The initial address list to be displayed in the dialog box, and whether to show only this address list.
    
- The number of recipient selectors, for example, whether to show all three labels of  **To**,  **Cc**, and  **Bcc**.
    
- The strings representing the title,  **To**,  **Cc**, and  **Bcc** labels where applicable. Long titles and labels will be truncated without resizing the width of the dialog box.
    
- Whether the user can select one or more address entries at a time.
    
- Whether to resolve recipient names before closing the dialog box.
    
- What to do if not all recipients are resolved.
    

To display the dialog box with names from an address list:


1. Use the  **[GetSelectNamesDialog](namespace-getselectnamesdialog-method-outlook.md)** method of the current session (indicated by **[Application.Session](application-session-property-outlook.md)**) to obtain an instance of the  **SelectNamesDialog** object for the current session.
    
2. Use the  **[AddressLists](namespace-addresslists-property-outlook.md)** property of the current session to obtain the collection of **AddressLists** for the current session.
    
3. By default, the dialog box is initialized with the address list that has  **[AddressList.IsInitialAddressList](addresslist-isinitialaddresslist-property-outlook.md)** set to **True**. If necessary, you can use  **[SelectNamesDialog.InitialAddressList](selectnamesdialog-initialaddresslist-property-outlook.md)** to initialize the dialog box with another **AddressList** from the **AddressLists** collection in Step 2.
    
4. Use  **[SelectNamesDialog.Display](selectnamesdialog-display-method-outlook.md)** to display the dialog box. This method returns a **True** or **False** depending on **[SelectNamesDialog.ForceResolution](selectnamesdialog-forceresolution-property-outlook.md)** and the user's response:
    
      - This method returns  **True** if **SelectNamesDialog.ForceResolution** is set, all selected names are resolved, and the user clicks **OK**.
    
  - It returns  **False** if **SelectNamesDialog.ForceResolution** is set, but not all the recipients are resolved.
    
  - It returns  **False** if **SelectNamesDialog.ForceResolution** is not set and the user clicks **OK**.
    
  - It returns  **False** if the user clicks **Cancel** or the **Close** icon.
    
5. If  **[SelectNamesDialog.Display](selectnamesdialog-display-method-outlook.md)** returns **True**, obtain the selected address entries using  **[SelectNamesDialog.Recipients](selectnamesdialog-recipients-property-outlook.md)**.
    


## See also


#### Concepts


 [How to: Identify the Global Address List or a Set of Address Lists with a Store](identify-the-global-address-list-or-a-set-of-address-lists-with-a-store.md)

