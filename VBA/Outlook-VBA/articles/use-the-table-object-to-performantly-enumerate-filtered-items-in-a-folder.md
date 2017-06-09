---
title: Use the Table Object to Performantly Enumerate Filtered Items in a Folder
ms.prod: outlook
ms.assetid: df82b04e-dffd-d621-10dd-34ee03df2051
ms.date: 06/08/2017
---


# Use the Table Object to Performantly Enumerate Filtered Items in a Folder

The code sample in this topic uses the  **[Table](table-object-outlook.md)** object to enumerate a filtered set of items in the Inbox that were last modified after May 1, 2005. For each of these items, the code sample prints these values: the subject, the time that the item was last modified, and whether the item is hidden. The procedure is as follows:


1. The sample defines a filter based on the value of the  **LastModificationTime** property of mail items.
    
2. It applies the filter to  **[Folder.GetTable](folder-gettable-method-outlook.md)** and obtains a **Table** of a subset of mail items in the Inbox that satisfies the filter.
    
     **Note**  The returned table contains a default set of properties for each of the filtered items:  **EntryID**,  **Subject**,  **CreationTime**,  **LastModificationTime**, and  **MessageClass**. 
3. It then uses  **[Columns.RemoveAll](columns-removeall-method-outlook.md)** and **[Columns.Add](columns-add-method-outlook.md)** to update the **Table** with the actually desired properties: **Subject**,  **LastModificationTime**, and the hidden attribute ( **PidTagAttributeHidden**). It specifies properties with their explicit built-in names if they exist (for example,  **Subject**,  **LastModificationTime**), and only when they don't, it references the properties by their namespaces (for example, the hidden attribute of a mail item).
    
     **Note**  The  **Table** objects returned from **Folder.GetTable** in Step 2 and **Columns.Add** in Step 3 contain different property values but for the same set of filtered items in the Inbox.
4. Lastly, it uses  **[Table.GetNextRow](table-getnextrow-method-outlook.md)** to enumerate the filtered items (until **[Table.EndOfTable](table-endoftable-property-outlook.md)** becomes true), displaying the values of the three desired properties for each item.
    






```vb
Sub DemoTable() 
 'Declarations 
 Dim Filter As String 
 Dim oRow As Outlook.Row 
 Dim oTable As Outlook.Table 
 Dim oFolder As Outlook.Folder 
 
 'Get a Folder object for the Inbox 
 Set oFolder = Application.Session.GetDefaultFolder(olFolderInbox) 
 
 'Define Filter to obtain items last modified after May 1, 2005 
 Filter = "[LastModificationTime] > '5/1/2005'" 
 'Restrict with Filter 
 Set oTable = oFolder.GetTable(Filter) 
 
 'Remove all columns in the default column set 
 oTable.Columns.RemoveAll 
 'Specify desired properties 
 With oTable.Columns 
 .Add ("Subject") 
 .Add ("LastModificationTime") 
 'PidTagAttributeHidden referenced by the MAPI proptag namespace 
 .Add ("http://schemas.microsoft.com/mapi/proptag/0x10F4000B") 
 End With 
 
 'Enumerate the table using test for EndOfTable 
 Do Until (oTable.EndOfTable) 
 Set oRow = oTable.GetNextRow() 
 Debug.Print (oRow("Subject")) 
 Debug.Print (oRow("LastModificationTime")) 
 Debug.Print (oRow("http://schemas.microsoft.com/mapi/proptag/0x10F4000B")) 
 Loop 
End Sub
```


