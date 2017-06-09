---
title: Conversation.GetTable Method (Outlook)
keywords: vbaol11.chm3382
f1_keywords:
- vbaol11.chm3382
ms.prod: outlook
api_name:
- Outlook.Conversation.GetTable
ms.assetid: 6c5a4ef5-c31d-6684-722a-f6f3b3fe6b55
ms.date: 06/08/2017
---


# Conversation.GetTable Method (Outlook)

Returns a  **[Table](table-object-outlook.md)** object that contains rows that represent all items in the conversation.


## Syntax

 _expression_ . **GetTable**

 _expression_ A variable that represents a **[Conversation](conversation-object-outlook.md)** object.


### Return Value

A  **Table** object that contains rows that represent all items in the conversation.


## Remarks

The  **GetTable** method returns a **Table** that has all items of the conversation as the rows. The default set of columns is shown in the following table.



|**Column**|**Property**|
|:-----|:-----|
|1| **EntryID**|
|2| **Subject**|
|3| **CreationTime**|
|4| **LastModificationTime**|
|5| **MessageClass**|
By default, the rows in the table are sorted by the  **ConversationIndex** property of the items.

To modify the default column set, use the  **[Add](columns-add-method-outlook.md)** , **[Remove](columns-remove-method-outlook.md)** , or **[RemoveAll](columns-removeall-method-outlook.md)** methods of the **[Columns](columns-object-outlook.md)** collection object.

The  **Table** object returned by this **GetTable** method does not include items in the conversation that have been moved to the Deleted Items folder.


## Example

The following Visual Basic for Applications (VBA) code example,  `DemoConversationTable`, assumes that there is a mail item opened in an inspector.  `DemoConversationTable` gets a **[Conversation](conversation-object-outlook.md)** object based on this mail item, and calls the **GetTable** method to get a **Table** of all the conversation items. To get specific information for each item in the conversation, which can span across stores, `DemoConversationTable` adds the store entry ID property, http://schemas.microsoft.com/mapi/proptag/0x0FFB0102, as a column to the table. As `DemoConversationTable` enumerates each item (represented by a row) in the table, it uses the store entry ID property that corresponds to that item to call the **[GetItemFromID](namespace-getitemfromid-method-outlook.md)** method of the **[NameSpace](namespace-object-outlook.md)** object to obtain the item object. The example then displays the subject and the number of attachments for that item.


 **Note**  Enumerating the conversation works only if the Outlook account is connected to a Microsoft Exchange Server that is running at least Microsoft Exchange Server 2010, or Outlook is running in cached mode against Microsoft Exchange Server 2007.


```vb
Sub DemoConversationTable() 
 Dim oConv As Outlook.Conversation 
 Dim oTable As Outlook.Table 
 Dim oRow As Outlook.Row 
 Dim oMail As Outlook.MailItem 
 Dim oItem As Outlook.MailItem 
 Const PR_STORE_ENTRYID As String = _ 
 "http://schemas.microsoft.com/mapi/proptag/0x0FFB0102" 
 
 On Error Resume Next 
 ' Obtain the current item for the active inspector. 
 Set oMail = Application.ActiveInspector.CurrentItem 
 
 If Not (oMail Is Nothing) Then 
 ' Obtain the Conversation object. 
 Set oConv = oMail.GetConversation 
 If Not (oConv Is Nothing) Then 
 Set oTable = oConv.GetTable 
 oTable.Columns.Add (PR_STORE_ENTRYID) 
 Do Until oTable.EndOfTable 
 Set oRow = oTable.GetNextRow 
 ' Use EntryID and StoreID to open the item. 
 Set oItem = Application.session.GetItemFromID( _ 
 oRow("EntryID"), _ 
 oRow.BinaryToString(PR_STORE_ENTRYID)) 
 Debug.Print oItem.Subject, _ 
 "Attachments.Count=" &; oItem.Attachments.count 
 Loop 
 End If 
 End If 
End Sub
```


## See also


#### Concepts


[Conversation Object](conversation-object-outlook.md)

