---
title: Page.Background Property (Publisher)
keywords: vbapb10.chm393249
f1_keywords:
- vbapb10.chm393249
ms.prod: publisher
api_name:
- Publisher.Page.Background
ms.assetid: 1bba32dc-0e7e-40ca-0f29-b67be6be518d
ms.date: 06/08/2017
---


# Page.Background Property (Publisher)

Sets or returns a  **PageBackground** object representing the background of the specified page.


## Syntax

 _expression_. **Background**

 _expression_A variable that represents a  **Page** object.


### Return Value

PageBackground


## Remarks

This property is for publication pages only. Any attempt to create a background for a master page will return a "Permission denied" error.


## Example

The following example creates a  **PageBackground** object and sets it to the background of the first page of the active document.


```vb
Dim objPageBackground As PageBackground 
Set objPageBackground = ActiveDocument.Pages(1).Background 
 
```


