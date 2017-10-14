---
title: Page.PageIndex Property (Publisher)
keywords: vbapb10.chm393224
f1_keywords:
- vbapb10.chm393224
ms.prod: publisher
api_name:
- Publisher.Page.PageIndex
ms.assetid: f64cc275-0474-7b97-d840-22e1e576d6f5
ms.date: 06/08/2017
---


# Page.PageIndex Property (Publisher)

Gets the index of the page in the  **Pages** collection of the active document. Read-only.


## Syntax

 _expression_. **PageIndex**

 _expression_A variable that represents a  **Page** object.


## Remarks

A  **PageIndex** property value is assigned to each page when it is added, and the value is incremented for each additional page. When pages are added or deleted, page index numbers are reassigned such that the first page is always 1 and the page index numbers always increment by 1. For example, in a publication with five pages, the page index numbers will be 1 through 5, regardless of the page number displayed on the pages themselves.


## Example

The following example displays the  **PageIndex**,  **PageNumber**, and  **PageID** properties for all the pages in the active publication.


```vb
Dim lngLoop As Long 
 
With ActiveDocument.Pages 
For lngLoop = 1 To .Count 
With .Item(lngLoop) 
Debug.Print "PageIndex = " &; .PageIndex _ 
&; " / PageNumber = " &; .PageNumber _ 
&; " / PageID = " &; .PageID 
End With 
Next lngLoop 
End With 
 

```


