---
title: Pages.Item Property (Publisher)
keywords: vbapb10.chm458752
f1_keywords:
- vbapb10.chm458752
ms.prod: publisher
api_name:
- Publisher.Pages.Item
ms.assetid: b220d9bf-321d-808a-d3cf-63a641e691e0
ms.date: 06/08/2017
---


# Pages.Item Property (Publisher)

Returns the specified  **[Page](page-object-publisher.md)** object from a **Pages** or **MasterPages** collection. Read-only.


## Syntax

 _expression_. **Item**( **_Item_**)

 _expression_A variable that represents a  **Pages** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Item|Required| **Long**|The number of the page to return. For  **MasterPages** collections, Item can either be 1 or 2 for the left and right master pages, respectively. For **Pages** collections, Item corresponds to a **Page** object's **[PageIndex](page-pageindex-property-publisher.md)** property.|

## Example

This example displays the page number, page index, and page ID of the first page in the active publication.


```vb
With ActiveDocument.Pages.Item(1) 
 Debug.Print "Page number = " &; .PageNumber 
 Debug.Print "Page index = " &; .PageIndex 
 Debug.Print "Page ID = " &; .PageID 
End With
```


