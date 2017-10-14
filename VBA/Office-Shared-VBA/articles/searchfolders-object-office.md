---
title: SearchFolders Object (Office)
keywords: vbaof11.chm258000
f1_keywords:
- vbaof11.chm258000
ms.prod: office
api_name:
- Office.SearchFolders
ms.assetid: 5958cafc-880e-ee9f-b2f5-be463bfe5232
ms.date: 06/08/2017
---


# SearchFolders Object (Office)

A collection of  **ScopeFolder** objects that determines which folders are searched.


## Remarks

For each application there is only a single  **SearchFolders** collection. The contents of the collection remains after the code that calls it has finished executing. Consequently, it is important to clear the collection unless you want to include folders from previous searches in your search.

You can use the  **Add** method of the **SearchFolders** collection to add a **ScopeFolder** object to the **SearchFolders** collection, however, it is usually simpler to use the **AddToSearchFolders** method of the **ScopeFolder** that you want to add, as there is only one **SearchFolders** collection for all searches.


## Methods



|**Name**|
|:-----|
|[Add](searchfolders-add-method-office.md)|
|[Remove](searchfolders-remove-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](searchfolders-application-property-office.md)|
|[Count](searchfolders-count-property-office.md)|
|[Creator](searchfolders-creator-property-office.md)|
|[Item](searchfolders-item-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
