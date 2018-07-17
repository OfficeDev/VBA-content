---
title: WebHiddenFields.Count Property (Publisher)
keywords: vbapb10.chm3997699
f1_keywords:
- vbapb10.chm3997699
ms.prod: publisher
api_name:
- Publisher.WebHiddenFields.Count
ms.assetid: 167c4c58-10cf-4dbb-5dfc-d60ab3856357
ms.date: 06/08/2017
---


# WebHiddenFields.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **WebHiddenFields** object.


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


