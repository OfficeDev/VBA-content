---
title: WebNavigationBarSets.Count Property (Publisher)
keywords: vbapb10.chm8454147
f1_keywords:
- vbapb10.chm8454147
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSets.Count
ms.assetid: ffe603c6-2c5a-de85-0924-aefa1dad269e
ms.date: 06/08/2017
---


# WebNavigationBarSets.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **WebNavigationBarSets** object.


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


