---
title: SimpleItems Object (Outlook)
keywords: vbaol11.chm3400
f1_keywords:
- vbaol11.chm3400
ms.prod: outlook
api_name:
- Outlook.SimpleItems
ms.assetid: b929ae28-fe5f-607e-37b5-ed6a304d4896
ms.date: 06/08/2017
---


# SimpleItems Object (Outlook)

Represents a set of possibly heterogeneous Microsoft Outlook items, with each member in the set tracking only a small, common set of properties that apply to Outlook items in general.


## Remarks

The  **SimpleItems** collection is used to represent child objects of a **[Conversation](conversation-object-outlook.md)** node object. This collection has only few members and serves the purpose of providing easy access to these items, as opposed to the **[Items](items-object-outlook.md)** and **[Results](results-object-outlook.md)** collections, which have more members.

The order of items in the collection is the same as the ordering of items in the conversation. The collection is ordered by the value of the  **CreationTime** property of each item in ascending order.


## Example

The following managed code is written in C#. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. You should use the following code in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.

The following code example assumes that the selected item in the explorer window is a mail item. The example obtains the conversation that the selected mail item is associated with, and enumerates each item in that conversation, displaying the subject of the item. The  `DemoConversation` method calls the **GetConversation** method of the selected mail item to obtain the associated **Conversation** object. `DemoConversation` then calls the **[GetTable](conversation-gettable-method-outlook.md)** and **[GetRootItems](conversation-getrootitems-method-outlook.md)** methods of the **Conversation** object to obtain a **[Table](table-object-outlook.md)** object and **[SimpleItems](simpleitems-object-outlook.md)** collection, respectively. `DemoConversation` calls the recurrent method `EnumerateConversation` to enumerate and display the subject of each item in that conversation.




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


## Methods



|**Name**|
|:-----|
|[Item](simpleitems-item-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](simpleitems-application-property-outlook.md)|
|[Class](simpleitems-class-property-outlook.md)|
|[Count](simpleitems-count-property-outlook.md)|
|[Parent](simpleitems-parent-property-outlook.md)|
|[Session](simpleitems-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
[How to: Obtain and Enumerate Selected Conversations](http://msdn.microsoft.com/library/3bba1e98-b2eb-c53d-354a-bdd899b65a59%28Office.15%29.aspx)
