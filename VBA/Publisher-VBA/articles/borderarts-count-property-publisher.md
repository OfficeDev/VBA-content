---
title: BorderArts.Count Property (Publisher)
keywords: vbapb10.chm7733251
f1_keywords:
- vbapb10.chm7733251
ms.prod: publisher
api_name:
- Publisher.BorderArts.Count
ms.assetid: 024cd14d-80f7-7372-c550-ef804661bbae
ms.date: 06/08/2017
---


# BorderArts.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **BorderArts** object.


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


## See also


#### Concepts


 [BorderArts Object](borderarts-object-publisher.md)

