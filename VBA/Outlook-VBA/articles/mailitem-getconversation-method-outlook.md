---
title: MailItem.GetConversation Method (Outlook)
keywords: vbaol11.chm3387
f1_keywords:
- vbaol11.chm3387
ms.prod: outlook
api_name:
- Outlook.MailItem.GetConversation
ms.assetid: f2017571-087c-1e83-4003-cb95097d43da
ms.date: 06/08/2017
---


# MailItem.GetConversation Method (Outlook)

Obtains a  **[Conversation](conversation-object-outlook.md)** object that represents the conversation to which this item belongs.


## Syntax

 _expression_ . **GetConversation**

 _expression_ A variable that represents a **[MailItem](mailitem-object-outlook.md)** object.


### Return Value

A  **Conversation** object that represents the conversation to which this item belongs.


## Remarks

 **GetConversation** returns **Null** ( **Nothing** in Visual Basic) if no conversation exists for the item. No conversation exists for an item in the following scenarios:


- The item has not been saved. An item can be saved programmatically, by user action, or by auto-save.
    
- For an item that can be sent (for example, a mail item, appointment item, or contact item), the item has not been sent.
    
- Conversations have been disabled through the Windows registry.
    
- The store does not support Conversation view (for example, Outlook is running in classic online mode against a version of Microsoft Exchange earlier than Microsoft Exchange Server 2010). Use the  **[IsConversationEnabled](store-isconversationenabled-property-outlook.md)** property of the **[Store](store-object-outlook.md)** object to determine whether the store supports Conversation view.
    



## Example

The following managed code is written in C#. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. You should use the following code in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.

The following code example assumes that the selected item in the explorer window is a mail item. The code example gets the conversation that the selected mail item is associated with, and enumerates each item in that conversation, displaying the subject of the item. The  `DemoConversation` method calls the **GetConversation** method of the selected mail item to get the associated **Conversation** object. `DemoConversation` then calls the **[GetTable](conversation-gettable-method-outlook.md)** and **[GetRootItems](conversation-getrootitems-method-outlook.md)** methods of the **Conversation** object to get a **[Table](table-object-outlook.md)** object and **[SimpleItems](simpleitems-object-outlook.md)** collection, respectively. `DemoConversation` calls the recurrent method `EnumerateConversation` to enumerate and display the subject of each item in that conversation.




```C#
void DemoConversation() 
{ 
 object selectedItem = 
 Application.ActiveExplorer().Selection[1]; 
 // This example uses only 
 // MailItem. Other item types such as 
 // MeetingItem and PostItem can participate 
 // in the conversation. 
 if (selectedItem is Outlook.MailItem) 
 { 
 // Cast selectedItem to MailItem. 
 Outlook.MailItem mailItem = 
 selectedItem as Outlook.MailItem; 
 // Determine the store of the mail item. 
 Outlook.Folder folder = mailItem.Parent 
 as Outlook.Folder; 
 Outlook.Store store = folder.Store; 
 if (store.IsConversationEnabled == true) 
 { 
 // Obtain a Conversation object. 
 Outlook.Conversation conv = 
 mailItem.GetConversation(); 
 // Check for null Conversation. 
 if (conv != null) 
 { 
 // Obtain Table that contains rows 
 // for each item in the conversation. 
 Outlook.Table table = conv.GetTable(); 
 Debug.WriteLine("Conversation Items Count: " + 
 table.GetRowCount().ToString()); 
 Debug.WriteLine("Conversation Items from Table:"); 
 while (!table.EndOfTable) 
 { 
 Outlook.Row nextRow = table.GetNextRow(); 
 Debug.WriteLine(nextRow["Subject"] 
 + " Modified: " 
 + nextRow["LastModificationTime"]); 
 } 
 Debug.WriteLine("Conversation Items from Root:"); 
 // Obtain root items and enumerate the conversation. 
 Outlook.SimpleItems simpleItems 
 = conv.GetRootItems(); 
 foreach (object item in simpleItems) 
 { 
 // In this example, enumerate only MailItem type. 
 // Other types such as PostItem or MeetingItem 
 // can appear in the conversation. 
 if (item is Outlook.MailItem) 
 { 
 Outlook.MailItem mail = item 
 as Outlook.MailItem; 
 Outlook.Folder inFolder = 
 mail.Parent as Outlook.Folder; 
 string msg = mail.Subject 
 + " in folder " + inFolder.Name; 
 Debug.WriteLine(msg); 
 } 
 // Call EnumerateConversation 
 // to access child nodes of root items. 
 EnumerateConversation(item, conv); 
 } 
 } 
 } 
 } 
} 
 
 
void EnumerateConversation(object item, 
 Outlook.Conversation conversation) 
{ 
 Outlook.SimpleItems items = 
 conversation.GetChildren(item); 
 if (items.Count > 0) 
 { 
 foreach (object myItem in items) 
 { 
 // In this example, enumerate only MailItem type. 
 // Other types such as PostItem or MeetingItem 
 // can appear in the conversation. 
 if (myItem is Outlook.MailItem) 
 { 
 Outlook.MailItem mailItem = 
 myItem as Outlook.MailItem; 
 Outlook.Folder inFolder = 
 mailItem.Parent as Outlook.Folder; 
 string msg = mailItem.Subject 
 + " in folder " + inFolder.Name; 
 Debug.WriteLine(msg); 
 } 
 // Continue recursion. 
 EnumerateConversation(myItem, conversation); 
 } 
 } 
} 
 

```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

