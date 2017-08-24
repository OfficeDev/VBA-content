---
title: Page.PageID Property (Publisher)
keywords: vbapb10.chm393223
f1_keywords:
- vbapb10.chm393223
ms.prod: publisher
api_name:
- Publisher.Page.PageID
ms.assetid: 07a87780-fb97-93ff-6f7d-1f1b72d3cb6a
ms.date: 06/08/2017
---


# Page.PageID Property (Publisher)

Returns a  **Long** indicating the unique identifier for a page in a publication. Read-only.


## Syntax

 _expression_. **PageID**

 _expression_A variable that represents a  **Page** object.


## Remarks

 **PageID** values are random numbers assigned to pages when they are added. These unique numbers do not change when pages are added or deleted. Also, these numbers do not start with 1, nor are they contiguous.


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


