---
title: Obtain and Enumerate Selected Conversations
ms.prod: outlook
ms.assetid: 3bba1e98-b2eb-c53d-354a-bdd899b65a59
ms.date: 06/08/2017
---


# Obtain and Enumerate Selected Conversations

By default, Microsoft Outlook displays items in the Inbox by conversation. If a user makes a selection in the Inbox, you can obtain the selection programmatically, including conversation headers and conversation items. The code example in this topic shows how to obtain a selection in the Inbox and enumerate the mail items in each conversation in the selection.

The example contains one method,  `DemoConversationHeadersFromSelection`. The method sets the current view to the Inbox, and then checks to see if the current view is a table view that displays conversations sorted by date. To obtain a selection, including any selected  [ConversationHeader](conversationheader-object-outlook.md) objects, `DemoConversationHeadersFromSelection` uses the [GetSelection](selection-getselection-method-outlook.md) method of the [Selection](selection-object-outlook.md) object, specifying the **OlSelectionContents.olConversationHeaders** constant as an argument. If conversation headers are selected, `DemoConversationHeadersFromSelection` uses the [SimpleItems](simpleitems-object-outlook.md) object to enumerate items in each selected conversation, and then displays the subject of those conversation items that are [MailItem](mailitem-object-outlook.md) objects.

The following managed code is written in C#. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. You should use the following code in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.




```C#
private void DemoConversationHeadersFromSelection() 
{ 
    // Obtain Inbox. 
    Outlook.Folder inbox = 
        Application.Session.GetDefaultFolder( 
        Outlook.OlDefaultFolders.olFolderInbox) 
        as Outlook.Folder; 
 
    // Set ActiveExplorer.CurrentFolder to Inbox. 
    // Inbox must be current folder. 
    Application.ActiveExplorer().CurrentFolder = inbox; 
 
    // Ensure that the current view is a table view. 
    if (inbox.CurrentView.ViewType == 
        Outlook.OlViewType.olTableView) 
    { 
        Outlook.TableView view = 
            inbox.CurrentView as Outlook.TableView; 
        // And check if the table view organizes conversations by date. 
        if (view.ShowConversationByDate == true) 
        { 
            Outlook.Selection selection = 
                Application.ActiveExplorer().Selection; 
            Debug.WriteLine("Selection.Count = " + selection.Count); 
             
             // Call GetSelection to create a Selection object 
            //  that includes any selected conversation header objects. 
            Outlook.Selection convHeaders = 
                selection.GetSelection( 
                Outlook.OlSelectionContents.olConversationHeaders) 
                as Outlook.Selection; 
            Debug.WriteLine("Selection.Count (ConversationHeaders) = "  
                + convHeaders.Count); 
 
            // Check if any conversation headers are selected. 
            if (convHeaders.Count >= 1) 
            { 
                foreach (Outlook.ConversationHeader convHeader in convHeaders) 
                { 
                    // Enumerate the items in each conversation header object. 
                    Outlook.SimpleItems items = convHeader.GetItems(); 
                    for (int i = 1; i <= items.Count; i++) 
                    { 
                        // Only enumerate MailItems in this example. 
                        if (items[i] is Outlook.MailItem) 
                        { 
                            Outlook.MailItem mail =  
                                items[i] as Outlook.MailItem; 
                            Debug.WriteLine(mail.Subject  
                                + " Received:" + mail.ReceivedTime); 
                        } 
                    } 
                } 
            } 
        } 
    } 
} 

```


