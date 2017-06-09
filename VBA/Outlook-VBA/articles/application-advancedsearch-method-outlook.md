---
title: Application.AdvancedSearch Method (Outlook)
keywords: vbaol11.chm728
f1_keywords:
- vbaol11.chm728
ms.prod: outlook
api_name:
- Outlook.Application.AdvancedSearch
ms.assetid: 7b433d8b-08b9-dff1-b854-287d76b47a90
ms.date: 06/08/2017
---


# Application.AdvancedSearch Method (Outlook)

Performs a search based on a specified DAV Searching and Locating (DASL) search string.


## Syntax

 _expression_ . **AdvancedSearch**( **_Scope_** , **_Filter_** , **_SearchSubFolders_** , **_Tag_** )

 _expression_ A variable that represents an **[Application](application-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Scope_|Required| **String**|The scope of the search. For example, the folder path of a folder. It is recommended that the folder path is enclosed within single quotes. Otherwise, the search might not return correct results if the folder path contains special characters including Unicode characters. To specify multiple folder paths, enclose each folder path in single quotes and separate the single quoted folder paths with a comma.|
| _Filter_|Optional| **Variant**|The DASL search filter that defines the parameters of the search.|
| _SearchSubFolders_|Optional| **Variant**|Determines if the search will include any of the folder's subfolders.|
| _Tag_|Optional| **Variant**|The name given as an identifier for the search.|

### Return Value

A  **[Search](search-object-outlook.md)** object that represents the results of the search.


## Remarks

You can run multiple searches simultaneously by calling the  **AdvancedSearch** method in successive lines of code. However, you should be aware that programmatically creating a large number of search folders can result in significant simultaneous search activity that would affect the performance of Outlook, especially if Outlook conducts the search in online Exchange mode.

The  **AdvancedSearch** method and related features in the Outlook object model do not create a Search Folder that will appear in the Outlook user interface. However, you can use the **[Save](search-save-method-outlook.md)** method of the **Search** object that is returned to create a Search Folder that will appear in the Search Folders list in the Outlook user interface.

Using the  _Scope_ parameter, you can specify one or more folders in the same store, but you may not specify multiple folders in multiple stores. To specify multiple folders in the same store for the _Scope_ parameter, use a comma character between each folder path and enclose each folder path in single quotes. For default folders such as Inbox or Sent Items, you can use the simple folder name instead of the full folder path. For example, the following two lines of code represent valid _Scope_ parameters:




```
Scope = "'Inbox', 'Sent Items'"
```




```
Scope = "'" &; Application.Session.GetDefaultFolder(olFolderInbox).FolderPath _  
    &; "','" &; Application.Session.GetDefaultFolder(olFolderSentMail).FolderPath &; "'"
```

The  _Filter_ parameter can be any valid DASL query. For additional information on DASL queries, see[Filtering Items](http://msdn.microsoft.com/library/4038e042-1b07-5d18-18b0-c2b58c9c42da%28Office.15%29.aspx) and[Referencing Properties by Namespace](http://msdn.microsoft.com/library/c1c7bfa9-64d7-81d2-84e7-f0a4c57780b3%28Office.15%29.aspx). Note that you cannot use a JET query for the  _Filter_ parameter of Advanced Search. If Instant Search is enabled on a store that contains a folder specified in the _Scope_ parameter, you can use Instant Search keywords to improve the performance of your search. If you use Instant Search keywords and Instant Search is not enabled, Outlook will return an error and your search will fail.


## Example

The following Visual Basic for Applications (VBA) example searches the  **Inbox** for items with subject equal to _Test_ and displays the names of the senders of the e-mail items returned by the search. The **[AdvancedSearchComplete](application-advancedsearchcomplete-event-outlook.md)** event procedure sets the boolean `blnSearchComp` to **True** when the search is complete. This boolean variable is used by the `TestAdvancedSearchComplete()` procedure to determine when the search is complete. The sample code must be placed in a class module such as `ThisOutlookSession`, and the  `TestAdvancedSearchComplete()` procedure must be called before the event procedure can be called by Outlook.


```vb
Public blnSearchComp As Boolean  
  
Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Search)  
    Debug.Print "The AdvancedSearchComplete Event fired"  
    If SearchObject.Tag = "Test" Then  
        m_SearchComplete = True  
    End If  
  
End Sub  
  
Sub TestAdvancedSearchComplete()  
    Dim sch As Outlook.Search  
    Dim rsts As Outlook.Results  
    Dim i As Integer  
    blnSearchComp = False  
    Const strF As String = "urn:schemas:mailheader:subject = 'Test'"  
    Const strS As String = "Inbox"     
    Set sch = Application.AdvancedSearch(strS, strF, ?Test?)   
    While blnSearchComp = False  
        DoEvents  
    Wend   
    Set rsts = sch.Results  
    For i = 1 To rsts.Count  
        Debug.Print rsts.Item(i).SenderName  
    Next  
End Sub
```

The following Microsoft Visual Basic for Applications example uses the  **AdvancedSearch** method to create a new search. The parameters of the search, as specified by the _Filter_ argument of the **AdvancedSearch** method, will return all items in the Inbox and Sent Items folders where the Subject phrase-matches or contains "Office". The user's Inbox and Sent Items folders are specified as the scope of the search and the **[SearchSubFolders](search-searchsubfolders-property-outlook.md)** property is set to **True** . When the search is complete, the **[GetTable](search-gettable-method-outlook.md)** method is called on the **[Search](search-object-outlook.md)** object for performant enumeration of search results.




```vb
Public m_SearchComplete As Boolean  
  
Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Search)  
    If SearchObject.Tag = "MySearch" Then  
        m_SearchComplete = True  
    End If  
End Sub  
  
Sub TestSearchForMultipleFolders()  
    Dim Scope As String  
    Dim Filter As String  
    Dim MySearch As Outlook.Search  
    Dim MyTable As Outlook.Table  
    Dim nextRow As Outlook.Row  
    m_SearchComplete = False  
    'Establish scope for multiple folders  
    Scope = "'" &; Application.Session.GetDefaultFolder( _  
    olFolderInbox).FolderPath _  
    &; "','" &; Application.Session.GetDefaultFolder( _  
    olFolderSentMail).FolderPath &; "'"  
    'Establish filter  
    If Application.Session.DefaultStore.IsInstantSearchEnabled Then  
        Filter = Chr(34) &; "urn:schemas:httpmail:subject" _  
        &; Chr(34) &; " ci_phrasematch 'Office'"  
    Else  
        Filter = Chr(34) &; "urn:schemas:httpmail:subject" _  
        &; Chr(34) &; " like '%Office%'"  
    End If  
    Set MySearch = Application.AdvancedSearch( _  
    Scope, Filter, True, "MySearch")  
    While m_SearchComplete <> True  
        DoEvents  
    Wend  
    Set MyTable = MySearch.GetTable  
    Do Until MyTable.EndOfTable  
        Set nextRow = MyTable.GetNextRow()  
        Debug.Print nextRow("Subject")  
    Loop  
End Sub
```


## See also


#### Concepts


[Application Object](application-object-outlook.md)

