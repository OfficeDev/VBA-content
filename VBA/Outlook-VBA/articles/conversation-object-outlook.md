---
title: Conversation Object (Outlook)
keywords: vbaol11.chm3388
f1_keywords:
- vbaol11.chm3388
ms.prod: outlook
api_name:
- Outlook.Conversation
ms.assetid: 2705d38a-ebc0-e5a7-208b-ffe1f5446b1b
ms.date: 06/08/2017
---


# Conversation Object (Outlook)

Represents a conversation that includes one or more items stored in one or more folders and stores.


## Remarks

The  **Conversation** object is an abstract, aggregated object. Although a conversation can include items of different types, the **Conversation** object does not correspond to a particular underlying MAPI **IMessage** object.

A conversation represents one or more items in one or more folders and stores. If you move an item in a conversation to the  **Deleted Items** folder and subsequently enumerate the conversation by using the **[GetChildren](http://msdn.microsoft.com/library/bc68ccd6-9d3c-c404-72b0-a21dbc99ed63%28Office.15%29.aspx)**, **[GetRootItems](http://msdn.microsoft.com/library/72c4d9fd-4f38-d081-7dc6-e9dbfad6d3aa%28Office.15%29.aspx)**, or **[GetTable](http://msdn.microsoft.com/library/6c5a4ef5-c31d-6684-722a-f6f3b3fe6b55%28Office.15%29.aspx)** method, the item will not be included in the returned object.

To obtain a  **Conversation** object for an existing conversation, use the **GetConversation** method of the item.

There are actions that you can apply to items in a conversation by calling the  **[SetAlwaysAssignCategories](http://msdn.microsoft.com/library/9b19f083-3aa9-8a0b-ea91-ff52fe46ad35%28Office.15%29.aspx)**, **[SetAlwaysDelete](http://msdn.microsoft.com/library/f13fce28-864e-a607-304d-a3722845cdd8%28Office.15%29.aspx)**, or **[SetAlwaysMoveToFolder](http://msdn.microsoft.com/library/52658b6d-c22c-a0e4-3743-4fe742bfbf9e%28Office.15%29.aspx)** method. Each of these actions is applied to all items in the conversation automatically when the method is called; the action is also applied to future items in the conversation as long as the action is still applicable to the conversation. There is no explicit save method on the **Conversation** object.

Also, when you apply an action to items in a conversation, the corresponding event occurs. For example, the  **[ItemChange](http://msdn.microsoft.com/library/6478357e-2a5a-300a-24e6-c125f8c81edd%28Office.15%29.aspx)** event of the **[Items](items-object-outlook.md)** object occurs when you call **SetAlwaysAssignCategories**, and the **[BeforeItemMove](http://msdn.microsoft.com/library/db75bc05-c80e-e6b8-d017-2150bc942712%28Office.15%29.aspx)** event of the **[Folder](folder-object-outlook.md)** object occurs when you call **SetAlwaysMoveToFolder**.


## Example

The following managed code is written in C#. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. You should use the following code in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.

The following code example assumes that the selected item in the explorer window is a mail item. The code example gets the conversation that the selected mail item is associated with, and enumerates each item in that conversation, displaying the subject of the item. The  `DemoConversation` method calls the **GetConversation** method of the selected mail item to get the associated **Conversation** object. `DemoConversation` then calls the **[GetTable](http://msdn.microsoft.com/library/6c5a4ef5-c31d-6684-722a-f6f3b3fe6b55%28Office.15%29.aspx)** and **[GetRootItems](http://msdn.microsoft.com/library/72c4d9fd-4f38-d081-7dc6-e9dbfad6d3aa%28Office.15%29.aspx)** methods of the **Conversation** object to get a **[Table](table-object-outlook.md)** object and **[SimpleItems](http://msdn.microsoft.com/library/b929ae28-fe5f-607e-37b5-ed6a304d4896%28Office.15%29.aspx)** collection, respectively. `DemoConversation` calls the recurrent method `EnumerateConversation` to enumerate and display the subject of each item in that conversation.




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
 // In this example, only enumerate MailItem type. 
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
 // In this example, only enumerate MailItem type. 
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
|[ClearAlwaysAssignCategories](http://msdn.microsoft.com/library/0494d8af-6569-c03d-99b1-be332c000985%28Office.15%29.aspx)|
|[GetAlwaysAssignCategories](http://msdn.microsoft.com/library/d09ae8ff-b725-cc09-9408-7b9039ee129f%28Office.15%29.aspx)|
|[GetAlwaysDelete](http://msdn.microsoft.com/library/95843bf3-7fff-fab0-ca7b-014ba290d718%28Office.15%29.aspx)|
|[GetAlwaysMoveToFolder](http://msdn.microsoft.com/library/ecad049d-338b-d5e0-f241-a9dddaeae316%28Office.15%29.aspx)|
|[GetChildren](http://msdn.microsoft.com/library/bc68ccd6-9d3c-c404-72b0-a21dbc99ed63%28Office.15%29.aspx)|
|[GetParent](http://msdn.microsoft.com/library/edcd31fb-f62e-4273-f827-ac1f704adc5e%28Office.15%29.aspx)|
|[GetRootItems](http://msdn.microsoft.com/library/72c4d9fd-4f38-d081-7dc6-e9dbfad6d3aa%28Office.15%29.aspx)|
|[GetTable](http://msdn.microsoft.com/library/6c5a4ef5-c31d-6684-722a-f6f3b3fe6b55%28Office.15%29.aspx)|
|[MarkAsRead](http://msdn.microsoft.com/library/94e764c8-e67a-0b8b-1f60-f892e3320e29%28Office.15%29.aspx)|
|[MarkAsUnread](http://msdn.microsoft.com/library/a8f580cb-a518-c5ca-778c-7d52ec22d2da%28Office.15%29.aspx)|
|[SetAlwaysAssignCategories](http://msdn.microsoft.com/library/9b19f083-3aa9-8a0b-ea91-ff52fe46ad35%28Office.15%29.aspx)|
|[SetAlwaysDelete](http://msdn.microsoft.com/library/f13fce28-864e-a607-304d-a3722845cdd8%28Office.15%29.aspx)|
|[SetAlwaysMoveToFolder](http://msdn.microsoft.com/library/52658b6d-c22c-a0e4-3743-4fe742bfbf9e%28Office.15%29.aspx)|
|[StopAlwaysDelete](http://msdn.microsoft.com/library/c759c9c8-bc43-ad5e-954c-88494c3dc4a6%28Office.15%29.aspx)|
|[StopAlwaysMoveToFolder](http://msdn.microsoft.com/library/3be830e9-ceea-369c-1f7b-966c68cfb8fd%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/d251a99d-96bc-e51b-02f0-fb61f2803f65%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/99e49411-5464-396e-09b9-28580179fdd1%28Office.15%29.aspx)|
|[ConversationID](http://msdn.microsoft.com/library/ee3cbe92-9e98-1151-1774-bd3884ab2aa3%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/e1b3f294-227a-27d9-84db-042da1be0caa%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/6f41faaa-e16a-d171-ed72-d2fef64a77f1%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
[Conversation Object Members](http://msdn.microsoft.com/library/09ff1e8e-7c5a-0b1e-e8e2-e259f66f71c8%28Office.15%29.aspx)
