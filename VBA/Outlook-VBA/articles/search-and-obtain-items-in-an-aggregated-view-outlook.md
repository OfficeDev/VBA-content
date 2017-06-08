---
title: Search and Obtain Items in an Aggregated View (Outlook)
ms.prod: outlook
ms.assetid: bd62f7b8-f110-ee0a-5930-877f14353a84
ms.date: 06/08/2017
---


# Search and Obtain Items in an Aggregated View (Outlook)

The  [GetTable](tableview-gettable-method-outlook.md) method of the [TableView](tableview-object-outlook.md) object can return a [Table](table-object-outlook.md) object that contains items from one or more folders in the same store or spanning over multiple stores, in an aggregated view. This is particularly useful when you want to access items returned from a search; for example, a search on all mail items in a store. This topic shows an example of how to use Instant Search to search for all items received from the manager of the current user that are marked important, and then display the subject of each search result.

The following code sample contains the  `GetItemsInView` method. `GetItemsInView` first makes a few checks to see if the current user of the Outlook session uses the Microsoft Exchange Server, whether the current user has a manager, and whether Instant Search is operational in the default store of the session. 

Because the eventual search relies on the [Search](explorer-search-method-outlook.md) method of the [Explorer](explorer-object-outlook.md) object, and the eventual result display uses the **GetTable** method, which is based on the current view of the current folder of the active explorer, `GetItemsInView` creates an explorer, displays the Inbox in this explorer, and sets up the search by using this **Explorer** object. `GetItemsInView` specifies the search criteria as important items from the current user's manager and the search scope as all folders of the [MailItem](mailitem-object-outlook.md) item type. 

After `GetItemsInView` calls the **Explorer.Search** method, any search results are displayed in this explorer, including items from other folders and stores that meet the search criteria. `GetItemsInView` obtains a **TableView** object that contains this explorer view of the search results. By using the **GetTable** method of this **TableView** object, `GetItemsInView` then obtains a **Table** object that contains the aggregated items returned from the search. Finally `GetItemsInView` displays the subject column of each row of the **Table** object that represents an item in the search results.

The following managed code is written in C#. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. 

You should use the following code in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.




```C#
private void GetItemsInView() 
{ 
    Outlook.AddressEntry currentUser = 
        Application.Session.CurrentUser.AddressEntry; 
 
    // Check if the current user uses the Exchange Server. 
    if (currentUser.Type == "EX") 
    { 
        Outlook.ExchangeUser manager = 
            currentUser.GetExchangeUser().GetExchangeUserManager(); 
 
        // Check if the current user has a manager. 
        if (manager != null) 
        { 
            string managerName = manager.Name; 
 
            // Check if Instant Search is enabled and operational in the default store. 
            if (Application.Session.DefaultStore.IsInstantSearchEnabled) 
            { 
                Outlook.Folder inbox = 
                    Application.Session.GetDefaultFolder( 
                    Outlook.OlDefaultFolders.olFolderInbox); 
 
                // Create a new explorer to display the Inbox as 
                // the current folder. 
                Outlook.Explorer explorer = 
                    Application.Explorers.Add(inbox, 
                    Outlook.OlFolderDisplayMode.olFolderDisplayNormal); 
 
                // Make the new explorer visible. 
                explorer.Display; 
 
                // Search for items from the manager marked important,  
                // from all folders of the same item type as the current folder,  
                // which is the MailItem item type. 
                string searchFor = 
                    "from:" + "\"" + managerName  
                    + "\"" + " importance:high"; 
                explorer.Search(searchFor, 
                    Outlook.OlSearchScope.olSearchScopeAllFolders); 
 
                // Any search results are displayed in that new explorer 
                // in an aggregated table view. 
                Outlook.TableView tableView =  
                    explorer.CurrentView as Outlook.TableView; 
 
                // Use GetTable of that table view to obtain items in that 
                // aggregated view in a Table object. 
                Outlook.Table table = tableView.GetTable(); 
                while (!table.EndOfTable) 
                { 
                    // Then display each row in the Table object 
                    // that represents an item in the search results. 
                    Outlook.Row nextRow = table.GetNextRow(); 
                    Debug.WriteLine(nextRow["Subject"]); 
                } 
            } 
        } 
    } 
} 

```


