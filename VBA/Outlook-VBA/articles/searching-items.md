---
title: Searching Items
ms.prod: OUTLOOK
ms.assetid: f0c24b9d-160e-3218-6979-2071a3135bfc
---


# Searching Items

This topic describes the entry points to search items in folders and access search results.



| **[Application.AdvancedSearch](application-advancedsearch-method-outlook.md)**|**[Explorer.Search](explorer-search-method-outlook.md)**|**[Search.GetTable](search-gettable-method-outlook.md)**|
|:-----|:-----|:-----|
|**Purpose**|Provides programmatic search on items in a specified folder based on a filter.|Performs a programmatic content indexer search that is analogous to a user executing a search from the Outlook user interface. |Provides an efficient way to access items (in a  ** [Table](table-object-outlook.md)**) returned by a prior  **Application.AdvancedSearch**. This entry point does not carry out a separate search.|
| **Scope of Search**|Folder specified as a search parameter.|Determined by the parameter  _SearchAllItems_. If  _SearchAllItems_ is True, the method will search across all folders that have the same folder type as the current folder (specified by the **[DefaultItemType](folder-defaultitemtype-property-outlook.md)** property of **[Explorer.CurrentFolder](explorer-currentfolder-property-outlook.md)**) and all stores that have been selected for search in the Search Options dialog box. If  _SearchAllItems_ is False, the method will search only the folder represented by **Explorer.CurrentFolder**.|Since the  **[Search](search-object-outlook.md)** object is returned from a prior **Application.AdvancedSearch** operation, the scope of the search associated with **Search.GetTable** is that of the prior **Application.AdvancedSearch**.|
| **Search Filter**|In DAV Searching and Locating (DASL) syntax.|Any valid keywords that are supported by Outlook search in the user interface. Search phrases are delimited by double quotes and can be concatenated to form a single search filter string.|Similar to the scope, the filter is of the search associated with  **Search.GetTable** is the filter parameter of the prior **Application.AdvancedSearch**.| **Search Completion**|Use the  **[AdvancedSearchComplete](application-advancedsearchcomplete-event-outlook.md)** event to determine when a given search has completed.|Does not provide a callback to indicate search completion.|Search is completed in the prior  **Application.AdvancedSearch**.  **Search.GetTable** only returns the search results.|
|**Search Results**|Access the search results by one of these means:|<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p><b><a href="search-results-property-outlook.md">Search.Results</a></b>    contains the search results as a <b><a href="results-object-outlook.md">Results</a></b> collection. Each item in the collection contain the full set of item properties.</p></li><li><p>Search.Save saves the results to a search folder. </p></li><li><p><b><a href="search-gettable-method-outlook.md">Search.GetTable</a></b> also returns the same set of items as in the <b>Results</b> collection, but each item  will contain only a default set of properties and therefore generally offers better performance.</p></li></ul> Search results are displayed in the Explorer for the current folder. To remove any search results in the Explorer, call  **[Explorer.ClearSearch](explorer-clearsearch-method-outlook.md)**. Search results are returned in a  **Table** which includes the same set of items returned from the prior **Application.AdvancedSearch**. Since the  **Table** only includes a limited set of properties per item, this is generally a more efficient way to access search results. To include properties other than the default in the search results, use **[Columns.Add](columns-add-method-outlook.md)** to get an updated **Table**. Since the item's Entry ID is one of the returned properties, you can also use  **[GetItemFromID](namespace-getitemfromid-method-outlook.md)** to obtain the item object, and access other item properties like **Body** and **AutoResolvedWinner** that are not supported by the **Table** object for performance reasons.|



