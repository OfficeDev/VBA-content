---
title: MailMergeFilters.Count Property (Publisher)
keywords: vbapb10.chm6750209
f1_keywords:
- vbapb10.chm6750209
ms.prod: publisher
api_name:
- Publisher.MailMergeFilters.Count
ms.assetid: 6ed658be-d3d0-ae5c-548d-ea724c9a8434
ms.date: 06/08/2017
---


# MailMergeFilters.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **MailMergeFilters** object.


## Example

This example displays the number of pages in the active document.


```vb
Sub CountNumberOfPages() 
 MsgBox "Your publication contains " &; _ 
 ActiveDocument.Pages.Count &; " page(s)." 
End Sub
```

This example displays the number of shapes in the active document.




```vb
Sub CountNumberOfShapes() 
 Dim intShapes As Integer 
 Dim pg As Page 
 
 For Each pg In ActiveDocument.Pages 
 intShapes = intShapes + pg.Shapes.Count 
 Next 
 
 MsgBox "Your publication contains " &; intShapes &; " shape(s)." 
End Sub
```


