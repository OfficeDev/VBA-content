---
title: Obtain Information for Multiple Accounts
ms.prod: outlook
ms.assetid: af587ee2-429a-252f-ecb6-2f058b9a37a8
ms.date: 06/08/2017
---


# Obtain Information for Multiple Accounts

Microsoft Outlook supports a profile that contains one or more accounts that are connected to a Microsoft Exchange Server. This topic shows how to obtain and display miscellaneous information about each account in the current profile.

The following method,  `EnumerateAccounts`, displays the account name, user name, and Simple Mail Transfer Protocol (SMTP) address for each account that is defined in the current profile. If the account is connected to an Exchange server,  `EnumerateAccounts` displays the Exchange server name and version information. And if the account resides on a delivery store, `EnumerateAccounts` displays the name of the default delivery store for the account.

 `EnumerateAccounts` accesses most of this information from the [Account](account-object-outlook.md) object, except when the **Account** object does not contain information about the user name and SMTP address. In that case, `EnumerateAccounts` uses the [AddressEntry](addressentry-object-outlook.md) and [ExchangeUser](exchangeuser-object-outlook.md) objects. `EnumerateAccounts` obtains the **AddressEntry** object by using the [AddressEntry](recipient-addressentry-property-outlook.md) property of the [Recipient](recipient-object-outlook.md) object obtained from the **[Account.CurrentUser](account-currentuser-property-outlook.md)** property. `EnumerateAccounts` obtains the **ExchangeUser** object by using the **[GetExchangeUser](addressentry-getexchangeuser-method-outlook.md)** method of the **AddressEntry** object. The following is the algorithm to obtain various information by using the **Account**,  **AddressEntry**, and  **ExchangeUser** objects:


- If the  **Account** object contains information about the user name and SMTP address, use the **Account** object to display the account name, user name, SMTP address, and Exchange server name and version information if the account is an Exchange account.
    
- Otherwise, the  **Account** object does not contain information about the user name and SMTP address, and proceed as follows:
    
      - If the account is not an Exchange account, use the  **AddressEntry** object to display the user name and SMTP address.
    
  - Otherwise, the account is an Exchange account, and proceed as follows:
    
      1. Use the  **Account** object to display the account name, Exchange server name, and Exchange version information.
    
      2. Use the  **ExchangeUser** object to display the user name and SMTP address.
    
The following managed code is written in C#. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. You should use the following code in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.



```C#
private void EnumerateAccounts() 
{ 
    Outlook.Accounts accounts = 
        Application.Session.Accounts; 
 
    // Enumerate each account defined in the current profile. 
    foreach (Outlook.Account account in accounts) 
    { 
        try 
        { 
            StringBuilder sb = new StringBuilder(); 
            sb.AppendLine("Account: " + account.DisplayName); 
 
            // If the account does not contain the SMTP address or 
            // user name, use the AddressEntry and ExchangeUser objects. 
            if (string.IsNullOrEmpty(account.SmtpAddress) 
                || string.IsNullOrEmpty(account.UserName)) 
            { 
                Outlook.AddressEntry oAE = 
                    account.CurrentUser.AddressEntry 
                    as Outlook.AddressEntry; 
 
                // If the account is an Exchange account, 
                // display also the Exchange server name and version. 
                if (oAE.Type == "EX") 
                { 
                    Outlook.ExchangeUser oEU = 
                        oAE.GetExchangeUser() 
                        as Outlook.ExchangeUser; 
  
                    // Use ExchangeUser object to display user name 
                    // and SMTP address. 
                    sb.AppendLine("UserName: " + 
                        oEU.Name); 
                    sb.AppendLine("SMTP: " + 
                        oEU.PrimarySmtpAddress); 
 
                    // Use Account object to display the Exchange 
                    // server name and version information. 
                    sb.AppendLine("Exchange Server: " + 
                        account.ExchangeMailboxServerName); 
                    sb.AppendLine("Exchange Server Version: " + 
                        account.ExchangeMailboxServerVersion);  
                } 
                // The account is not connected to an Exchange 
                // Server, use the AddressEntry object to display only  
                // the user name and SMTP address. 
                else 
                { 
                    sb.AppendLine("UserName: " + 
                        oAE.Name); 
                    sb.AppendLine("SMTP: " + 
                        oAE.Address); 
                } 
            } 
            // The account contains SMTP address and 
            // user name,  then the Account object is sufficient.  
            else 
            { 
                sb.AppendLine("UserName: " + 
                    account.UserName); 
                sb.AppendLine("SMTP: " + 
                    account.SmtpAddress); 
 
                // If the account is an Exchange account, 
                // display also the Exchange server name and version. 
                if(account.AccountType ==  
                    Outlook.OlAccountType.olExchange) 
                { 
                    sb.AppendLine("Exchange Server: " + 
                        account.ExchangeMailboxServerName); 
                    sb.AppendLine("Exchange Server Version: " + 
                        account.ExchangeMailboxServerVersion);  
                } 
            } 
 
            // If the account is connected to a delivery store, 
            // display the store name as well. 
            if(account.DeliveryStore !=null) 
            { 
                sb.AppendLine("Delivery Store: " + 
                    account.DeliveryStore.DisplayName); 
            } 
            sb.AppendLine("---------------------------------"); 
            Debug.Write(sb.ToString()); 
        } 
        catch (Exception ex) 
        { 
            Debug.WriteLine(ex.Message); 
        } 
    } 
} 

```


