---
title: Search.IsSynchronous Property (Outlook)
keywords: vbaol11.chm2254
f1_keywords:
- vbaol11.chm2254
ms.prod: outlook
api_name:
- Outlook.Search.IsSynchronous
ms.assetid: e240cc55-26c3-a560-4ee2-84b15da95e52
ms.date: 06/08/2017
---


# Search.IsSynchronous Property (Outlook)

Returns a  **Boolean** indicating whether the search is synchronous. Read-only.


## Syntax

 _expression_ . **IsSynchronous**

 _expression_ A variable that represents a **Search** object.


## Remarks

A search can be synchronous or asynchronous. If the search is synchronous, code execution will pause until the search has completed. Conversely, if the search is asynchronous, code execution will continue even though the search has not completed. In this case, use the  **[Search](search-object-outlook.md)** object's **[Stop](search-stop-method-outlook.md)** method to halt the search. In order to get meaningful results from an asynchronous search, use the **[AdvancedSearchComplete](application-advancedsearchcomplete-event-outlook.md)** event to notify you when the search has finished.


## See also


#### Concepts


[Search Object](search-object-outlook.md)

