---
title: Search the Inbox for Items with Subject Containing "Office"
ms.prod: outlook
ms.assetid: 2a2fa978-8652-edd4-ad8f-efeffc8faf65
ms.date: 06/08/2017
---


# Search the Inbox for Items with Subject Containing "Office"

This topic shows two code samples that use DASL queries to search for items in the Inbox that contain "Office" in the subject line. The first code sample uses  **[Folder.GetTable](folder-gettable-method-outlook.md)** and the second uses **[Application.AdvancedSearch](application-advancedsearch-method-outlook.md)** to apply the DASL query.

Each of the code samples uses the content indexer keyword  **ci_phrasematch** in a DASL filter on the property **http://schemas.microsoft.com/mapi/proptag/0x0037001E** (the **Subject** property referenced by the MAPI ID namespace) to search for the word "office" in the subject. It applies the filter to items in the Inbox (by using **Folder.GetTable** or **Application.AdvancedSearch**), and prints the subject line of each item returned from the search.

 **Note**  The match is not case-sensitive so any item containing "Office" or "office" in the subject will be returned by  **Folder.GetTable** or **Application.AdvancedSearch**. Notice that each sample prints the subject of each row in the resultant  **[Table](table-object-outlook.md)**. It chooses to use the lighter weight  **Table** object instead of the **[Search.Results](search-results-property-outlook.md)** object for better performance. The **Subject** property is included in a **Table** returned by a search on any folder. 
 
 But like any folder in Outlook, the Inbox can contain heterogenous items and is not confined to mail items. If you want to access a property that is specific to a certain item type in the Inbox, use **[Columns.Add](columns-add-method-outlook.md)** to include that property and update the **Table**, and for each row returned in the  **Table**, check the message type of the item before accessing the property.

This code sample uses  **Folder.GetTable** to do the search:



```vb
Sub RestrictTableForInbox() 
    Dim oT As Outlook.Table 
    Dim strFilter As String 
    Dim oRow As Outlook.Row 
     
    'Construct filter for Subject containing 'Office' 
    Const PropTag  As String = "http://schemas.microsoft.com/mapi/proptag/" 
    strFilter = "@SQL=" &; Chr(34) &; PropTag  _ 
        &; "0x0037001E" &; Chr(34) &; " ci_phrasematch 'Office'" 
     
    'Do search and obtain Table on Inbox 
    Set oT = Application.Session.GetDefaultFolder(olFolderInbox).GetTable(strFilter) 
     
    'Print Subject of each returned item 
    Do Until oT.EndOfTable 
        Set oRow = oT.GetNextRow 
        Debug.Print oRow("Subject") 
    Loop 
End Sub
```

This code sample uses  **Application.AdvancedSearch** to do the search:



```vb
Public blnSearchComp As Boolean 
Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Search) 
    MsgBox "The AdvancedSearchComplete Event fired" 
    blnSearchComp = True 
End Sub 
 
Sub TestSearchWithTable() 
    Dim oSearch As Search 
    Dim oTable As Table 
    Dim strQuery As String 
    Dim oRow As Row 
         
    blnSearchComp = False 
     
    'Construct filter. 0x0037001E represents Subject 
    strQuery = _ 
        "http://schemas.microsoft.com/mapi/proptag/0x0037001E" &; _ 
        " ci_phrasematch 'Office'" 
     
    'Do search 
    Set oSearch = _ 
        Application.AdvancedSearch("Inbox", strQuery, False, "Test") 
    While blnSearchComp = False 
        DoEvents 
    Wend 
 
    'Obtain Table 
    Set oTable = oSearch.GetTable 
     
    'Print Subject of each returned item 
    Do Until oTable.EndOfTable 
        Set oRow = oTable.GetNextRow 
        Debug.Print oRow("Subject") 
    Loop 
End Sub
```


