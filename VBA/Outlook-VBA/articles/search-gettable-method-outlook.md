---
title: Search.GetTable Method (Outlook)
keywords: vbaol11.chm2261
f1_keywords:
- vbaol11.chm2261
ms.prod: outlook
api_name:
- Outlook.Search.GetTable
ms.assetid: 3aba6b77-73a3-9620-9c18-b2e03c7b63bc
ms.date: 06/08/2017
---


# Search.GetTable Method (Outlook)

Obtains a  **[Table](table-object-outlook.md)** object that contains items filtered by the _Filter_ parameter in a preceding **[Application.AdvancedSearch](application-advancedsearch-method-outlook.md)** method call.


## Syntax

 _expression_ . **GetTable**

 _expression_ A variable that represents a **Search** object.


### Return Value

A  **Table** that contains items that meet the criteria specified by the _Filter_ parameter in a preceding **Application.AdvancedSearch** method call.


## Remarks

Unlike  **[Folder.GetTable](folder-gettable-method-outlook.md)** , **Search.GetTable** does not accept a _Filter_ parameter. The filter for the **Table** is determined by **[Search.Filter](search-filter-property-outlook.md)** . Since **Search.Filter** is a read-only property, the _Filter_ parameter for the **Application.AdvancedSearch** method establishes the filter for the **Table** object returned by **Search.GetTable** .

The  _Filter_ parameter supplied to **Application.AdvancedSearch** must be a DASL query. The filter for **AdvancedSearch** will not accept a JET query. Do not prefix a DASL query for **AdvancedSearch** with "@SQL=". If you add the "@SQL=" prefix, your query will raise an error. For more information on filters, see[Filtering Items](http://msdn.microsoft.com/library/4038e042-1b07-5d18-18b0-c2b58c9c42da%28Office.15%29.aspx).

 **Search.GetTable** returns a **Table** with the default column set for the folder type of the parent **Folder** . To modify the default column set, use the **[Add](columns-add-method-outlook.md)** , **[Remove](columns-remove-method-outlook.md)** , or **[RemoveAll](columns-removeall-method-outlook.md)** methods of the **[Columns](columns-object-outlook.md)** collection object. For more information on default column sets, see[Default Properties Displayed in a Table Object](http://msdn.microsoft.com/library/649c64f3-2d1e-23f1-bf13-3368da79e62b%28Office.15%29.aspx).

Unlike  **Folder.GetTable** , you cannot use **[Table.Restrict](table-restrict-method-outlook.md)** to apply subsequent filters to a **Table** that is based on the **Search** object. Specify a new filter in **Application.AdvancedSearch** to re-apply a filter.


## See also


#### Concepts


[Search Object](search-object-outlook.md)

