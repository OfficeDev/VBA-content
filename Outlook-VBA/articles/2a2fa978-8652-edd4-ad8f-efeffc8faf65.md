
# Search the Inbox for Items with Subject Containing "Office"

 **Last modified:** July 28, 2015

 _**Applies to:** Outlook 2013_

This topic shows two code samples that use DASL queries to search for items in the Inbox that contain "Office" in the subject line. The first code sample uses  ** [Folder.GetTable](08d184cb-0c41-01b1-abc5-305476380f8b.md)** and the second uses ** [Application.AdvancedSearch](7b433d8b-08b9-dff1-b854-287d76b47a90.md)** to apply the DASL query.

Each of the code samples uses the content indexer keyword  **ci_phrasematch** in a DASL filter on the property **http://schemas.microsoft.com/mapi/proptag/0x0037001E** (the **Subject** property referenced by the MAPI ID namespace) to search for the word "office" in the subject. It applies the filter to items in the Inbox (by using **Folder.GetTable** or **Application.AdvancedSearch**), and prints the subject line of each item returned from the search.

 **Note**  The match is not case-sensitive so any item containing "Office" or "office" in the subject will be returned by  **Folder.GetTable** or **Application.AdvancedSearch**. Notice that each sample prints the subject of each row in the resultant  ** [Table](0affaafd-93fe-227a-acee-e09a86cadc20.md)**. It chooses to use the lighter weight  **Table** object instead of the ** [Search.Results](405166fa-d0bc-33d2-f4aa-908fb821edd6.md)** object for better performance. The **Subject** property is included in a **Table** returned by a search on any folder. But like any folder in Outlook, the Inbox can contain heterogenous items and is not confined to mail items. If you want to access a property that is specific to a certain item type in the Inbox, use ** [Columns.Add](d438cfeb-629f-4234-6f4f-ffa086ef9a41.md)** to include that property and update the **Table**, and for each row returned in the  **Table**, check the message type of the item before accessing the property.

This code sample uses  **Folder.GetTable** to do the search:



```
Sub RestrictTableForInbox() 
    Dim oT As Outlook.Table 
    Dim strFilter As String 
    Dim oRow As Outlook.Row 
     
    'Construct filter for Subject containing 'Office' 
    Const PropTag  As String = "http://schemas.microsoft.com/mapi/proptag/" 
    strFilter = "@SQL=" &amp; Chr(34) &amp; PropTag  _ 
        &amp; "0x0037001E" &amp; Chr(34) &amp; " ci_phrasematch 'Office'" 
     
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



```
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
        "http://schemas.microsoft.com/mapi/proptag/0x0037001E" &amp; _ 
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

