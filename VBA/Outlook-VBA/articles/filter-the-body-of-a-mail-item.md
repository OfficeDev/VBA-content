---
title: Filter the Body of a Mail Item
ms.prod: outlook
ms.assetid: 15d8fec5-4b3d-340b-2394-479abf29847c
ms.date: 06/08/2017
---


# Filter the Body of a Mail Item

The code sample in this topic shows how to use content indexing in a DASL query to search for mail items that contain a specific word in the body. 

The code sample sets up a DASL filter on the property  **urn:schemas:httpmail:textdescription** (the **Body** property referenced by its DAV namespace) and uses the content indexer keyword **ci_phrasematch** to search for the word "office" in the body. It then applies the filter to items in the current folder. To access the filter results, it uses the **[Table](table-object-outlook.md)** object and prints the subject line of each item.

Notice that this sample prints the subject of each row in the returned  **Table**; the  **Subject** property is included in a **Table** returned by a search on any folder. But generally, a folder in Outlook can contain heterogenous items and is not confined to a single item type. If you want to access a property that is specific to an item type, use **[Columns.Add](columns-add-method-outlook.md)** to include that property and update the **Table**, and for each row returned in the  **Table**, check the message type of the item before accessing the property.


 **Note**  Content indexing in a DASL query provides better performance than the  **like** keyword. However, you can filter only on the text of the item body; if the body contains HTML tags, as in an HTML-formatted mail item, the tags will not be filtered. The match is not case-sensitive, so for example, any item containing "Office" or "office" in the body will be returned by **[Folder.GetTable](folder-gettable-method-outlook.md)**. You can also return up to the first 255 characters of the body in a column of a Table, by adding the column (denoted by  **urn:schemas:httpmail:textdescription**) to the  **Table**. You cannot use a Jet query to filter on the  **Body** property.




```vb
Sub RestrictUsingBody() 
 Dim strFilter As String 
 Dim oT As Table 
 Dim oRow as Row 
 
 'Create DASL query for Body using content indexing phrase match for 'office' 
 strFilter = "@SQL=" &; Chr(34) &; "urn:schemas:httpmail:textdescription" _ 
 &; Chr(34) &; " ci_phrasematch 'office'" 
 'Obtain Table by applying the filter on the current folder 
 Set oT = Application.ActiveExplorer.CurrentFolder.GetTable(strFilter) 
 'Print subject line of each mail item in current folder that contains 'office' in the body 
 Do Until oT.EndOfTable 
 Set oRow = oT.GetNextRow 
 Debug.Print oRow("Subject") 
 Loop 
End Sub
```


