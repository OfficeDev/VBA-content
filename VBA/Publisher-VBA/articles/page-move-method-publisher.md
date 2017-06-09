---
title: Page.Move Method (Publisher)
keywords: vbapb10.chm393250
f1_keywords:
- vbapb10.chm393250
ms.prod: publisher
api_name:
- Publisher.Page.Move
ms.assetid: 754cfe41-0853-a2cf-59ee-85db68fb871a
ms.date: 06/08/2017
---


# Page.Move Method (Publisher)

Moves the specified page to the specified index in the  **Pages** collection.


## Syntax

 _expression_. **Move**( **_Page_**,  **_After_**)

 _expression_A variable that represents a  **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Page|Required| **Long**|The index number of the  **Pages** collection where the specified page will be moved.|
|After|Optional| **Boolean**| **True** if the page will be inserted after the specified index number of the Pages collection specified by the Page parameter. Deafult is **True**.|

## Example

This example moves the first page of the publication before the third page of the publication. This example assumes that there are at least three pages in the document.


```vb
ActiveDocument.Pages(1).Move page:=3, After:=False
```


