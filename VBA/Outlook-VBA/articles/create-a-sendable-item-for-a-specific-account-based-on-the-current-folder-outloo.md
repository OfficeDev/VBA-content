---
title: Create a Sendable Item for a Specific Account Based on the Current Folder (Outlook)
ms.prod: outlook
ms.assetid: 758e2e9c-3633-2e77-b9e0-14bb8078cf0b
ms.date: 06/08/2017
---


# Create a Sendable Item for a Specific Account Based on the Current Folder (Outlook)

When you use the  [CreateItem](application-createitem-method-outlook.md) method of the [Application](application-object-outlook.md) object to create a Microsoft Outlook item, the item is created for the primary account for that session. In a session where multiple accounts are defined in the profile, you can create an item for a specific IMAP, POP, or Microsoft Exchange account. If there are multiple accounts in the current profile and you create a sendable item in the user interface, for example, by clicking **New E-mail** or **New Meeting**, an inspector displays a new mail item or meeting request in compose mode, and then you can select the account from which to send the item. This topic shows how to programmatically create a sendable item and send it by using a specific sending account. The topic has two code samples that show how to create a  [MailItem](mailitem-object-outlook.md) and an [AppointmentItem](appointmentitem-object-outlook.md) for a specific account that is determined by the current folder in the active explorer.

The following managed code is written in C#. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. You should use the following code in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.

The first method below,  `CreateMailItemFromAccount`, creates a  **MailItem** for a specific account and displays it in compose mode; the default delivery store of the specific account is the same as the store for the folder that is displayed in the active explorer. The current user of the account is set as the sender. `CreateMailItemFromAccount` first identifies the appropriate account by matching the store of the current folder (obtained from the **[Folder.Store](folder-store-property-outlook.md)** property) with the default delivery store of each account (obtained with the **[Account.DeliveryStore](account-deliverystore-property-outlook.md)** property) that is defined in the **[Accounts](accounts-object-outlook.md)** collection for the session. `CreateMailItemFromAccount` then creates the **MailItem**. To associate the item with the account,  `CreateMailItemFromAccount` assigns the user of the account as the sender of the item by setting the [AddressEntry](addressentry-object-outlook.md) object for the account's user to the [Sender](mailitem-sender-property-outlook.md) property of the **MailItem**. Assigning the  **Sender** property is the important step because otherwise, the **MailItem** is created for the primary account. At the end of the method, `CreateMailItemFromAccount` displays the **MailItem**. Note that if the current folder is not on a delivery store,  `CreateMailItemFromAccount` simply creates the **MailItem** for the primary account for the session.




```C#
private void CreateMailItemFromAccount() 
{ 
    Outlook.AddressEntry addrEntry = null; 
 
    // Get the store for the current folder. 
    Outlook.Folder folder = 
        Application.ActiveExplorer().CurrentFolder  
        as Outlook.Folder; 
    Outlook.Store store = folder.Store; 
     
    Outlook.Accounts accounts = 
        Application.Session.Accounts; 
 
    // Match the delivery store of each account with the  
    // store for the current folder. 
    foreach (Outlook.Account account in accounts) 
    { 
        if (account.DeliveryStore.StoreID ==  
            store.StoreID) 
        { 
            addrEntry = 
                account.CurrentUser.AddressEntry; 
            break; 
        } 
    } 
 
    // Create MailItem. Account is either the primary 
    // account or the account with a delivery store 
    // that matches the store for the current folder. 
    Outlook.MailItem mail = 
        Application.CreateItem( 
        Outlook.OlItemType.olMailItem) 
        as Outlook.MailItem; 
 
    if (addrEntry != null) 
    { 
        //Set Sender property. 
        mail.Sender = addrEntry; 
        mail.Display(false); 
    } 
} 

```

The next method,  `CreateMeetingRequestFromAccount`, is similar to  `CreateMailItemFromAccount` except that it creates an **AppointmentItem** instead of a **MailItem**, and associates the  **AppointmentItem** with the account by using its [SendUsingAccount](appointmentitem-sendusingaccount-property-outlook.md) property. `CreateMeetingRequestFromAccount` creates an **AppointmentItem** in the Calender folder of an account whose default delivery store is the same as the store for the folder that is displayed in the active explorer. `CreateMeetingRequestFromAccount` first identifies the appropriate account by matching the store of the current folder (obtained from the **Folder.Store** property) with the default delivery store of each account (otained with the **Account.DeliveryStore** property) that is defined in the **Accounts** collection for the session. `CreateMeetingRequestFromAccount` then creates the **AppointmentItem**. To associate the item with the account,  `CreateMeetingRequestFromAccount` assigns that account as the item's sending account by setting the [Account](account-object-outlook.md) object to the **SendUsingAccount** property of the **AppointmentItem**. Assigning the  **SendUsingAccount** property is the important step because otherwise, the **AppointmentItem** is created for the primary account. At the end of the method, `CreateMeetingRequestFromAccount` displays the **AppointmentItem**. Note that if the current folder is not on a delivery store,  `CreateMeetingRequestFromAccount` simply creates the **AppointmentItem** for the primary account for the session.



```C#
private void CreateMeetingRequestFromAccount() 
{ 
    Outlook.Account acct = null; 
 
    // Get the store for the current folder. 
    Outlook.Folder folder = 
        Application.ActiveExplorer().CurrentFolder 
        as Outlook.Folder; 
    Outlook.Store store = folder.Store; 
 
    Outlook.Accounts accounts = 
        Application.Session.Accounts; 
 
    // Match the delivery store of each account with the  
    // store for the current folder. 
    foreach (Outlook.Account account in accounts) 
    { 
        if (account.DeliveryStore.StoreID == 
            store.StoreID) 
        { 
            acct = account; 
            break; 
        } 
    } 
  
    // Create AppointmentItem. Account is either the primary 
    // account or the account with a delivery store 
    // that matches the store for the current folder. 
    Outlook.AppointmentItem appt = 
        Application.CreateItem( 
        Outlook.OlItemType.olAppointmentItem) 
        as Outlook.AppointmentItem; 
 
    appt.MeetingStatus =  
        Outlook.OlMeetingStatus.olMeeting; 
    if (acct != null) 
    { 
        //Set SendUsingAccount property. 
        appt.SendUsingAccount=acct; 
        appt.Display(false); 
    } 
} 

```


