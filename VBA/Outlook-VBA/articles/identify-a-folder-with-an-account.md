---
title: Identify a Folder with an Account
ms.prod: outlook
ms.assetid: 64dfbe81-933a-0929-e18c-a927156e50d4
ms.date: 06/08/2017
---


# Identify a Folder with an Account

In a Microsoft Outlook session that has multiple accounts defined in the profile, the folder that is displayed in the active explorer does not necessarily reside on the default store for that session; instead, it can reside on one of the multiple stores associated with the multiple accounts. This topic shows how to identify the account whose default delivery store is the same store that hosts the folder.

In the following code sample, the  `DisplayAccountForCurrentFolder` function calls the `GetAccountForFolder` function to identify the account whose default delivery store hosts the current folder, and then displays the name of the folder. `GetAccountForFolder` matches the store of the current folder (obtained from the **[Folder.Store](folder-store-property-outlook.md)** property) with the default delivery store of each account (obtained with the **[Account.DeliveryStore](account-deliverystore-property-outlook.md)** property) that is defined in the [Accounts](accounts-object-outlook.md) collection for the session. `GetAccountForFolder` returns the [Account](account-object-outlook.md) object when a match is found; otherwise, it returns null.

The following managed code is written in C#. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. You should use the following code in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.




```C#
private void DisplayAccountForCurrentFolder() 
{ 
    Outlook.Folder myFolder = Application.ActiveExplorer().CurrentFolder  
        as Outlook.Folder; 
    string msg = "Account for Current Folder:" + "\n" + 
        GetAccountForFolder(myFolder).DisplayName; 
    MessageBox.Show(msg, 
        "GetAccountForFolder", 
        MessageBoxButtons.OK, 
        MessageBoxIcon.Information); 
} 
 
Outlook.Account GetAccountForFolder(Outlook.Folder folder) 
{ 
    // Obtain the store on which the folder resides. 
    Outlook.Store store = folder.Store; 
 
    // Enumerate the accounts defined for the session. 
    foreach (Outlook.Account account in Application.Session.Accounts) 
    { 
        // Match the DefaultStore.StoreID of the account 
        // with the Store.StoreID for the currect folder. 
        if (account.DeliveryStore.StoreID  == store.StoreID) 
        { 
            // Return the account whose default delivery store 
            // matches the store of the given folder. 
            return account; 
        } 
     } 
     // No account matches, so return null. 
     return null; 
}
```


