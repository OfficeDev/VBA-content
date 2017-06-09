---
title: Pages.Add Method (Publisher)
keywords: vbapb10.chm458757
f1_keywords:
- vbapb10.chm458757
ms.prod: publisher
api_name:
- Publisher.Pages.Add
ms.assetid: 3c22aa15-c1dc-94c8-62d6-a1bc9635cd89
ms.date: 06/08/2017
---


# Pages.Add Method (Publisher)

Adds a new  **Page** object to the specified **Pages** object and returns the new **Page** object.


## Syntax

 _expression_. **Add**( **_Count_**,  **_After_**,  **_DuplicateObjectsOnPage_**,  **_AddHyperlinkToWebNavBar_**)

 _expression_A variable that represents a  **Pages** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Count|Required| **Long**|The number of new pages to add.|
|After|Required| **Long**|The page index of the page after which to add the new pages. A zero for this argument adds new pages at the beginning of the publication.|
|DuplicateObjectsOnPage|Optional| **Long**|The page index of the page from which objects should be copied to the new pages. If this argument is omitted, the new pages will be blank. The default is -1: do not duplicate pages.|
|AddHyperlinkToWebNavBar|Optional| **Boolean**|Specifies whether links to the new pages will be added to the automatic navigation bars of existing pages. If  **True**, links to the new pages will be added to the automatic navigation bars of existing pages only. If  **False**, links to the new pages will not be added to the automatic navigation bars of existing pages or new pages added in the future. Default is  **False**.|

### Return Value

Page


## Example

The following example adds four new pages after the first page in the publication and copies all the objects from the first page to the new pages.


```vb
Dim pgNew As Page 
 
Set pgNew = ActiveDocument.Pages _ 
 .Add(Count:=4, After:=1, DuplicateObjectsOnPage:=1)
```

The following example demonstrates adding two new pages to the publication and setting the  **AddHyperlinkToWebNavBar** parameter to **True** for these two pages. This specifies that links to these two new pages be added to the automatic navigation bars of existing pages and those added in the future.

Another page is then added to the publication, and the  **AddHyperlinkToWebNavBar** is omitted. This means that the **IncludePageOnNewWebNavigationBars** property is **False** for the newly added page, and links to this page will not be included in the automatic navigation bars of existing pages.




```vb
Dim thePage As page 
Dim thePage2 As page 
 
Set thePage = ActiveDocument.Pages.Add(Count:=2, _ 
 After:=4, AddHyperlinkToWebNavBar:=True) 
 
Set thePage2 = ActiveDocument.Pages.Add(Count:=1, After:=6)
```


