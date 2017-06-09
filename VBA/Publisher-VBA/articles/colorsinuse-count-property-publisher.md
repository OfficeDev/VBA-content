---
title: ColorsInUse.Count Property (Publisher)
keywords: vbapb10.chm2949122
f1_keywords:
- vbapb10.chm2949122
ms.prod: publisher
api_name:
- Publisher.ColorsInUse.Count
ms.assetid: 2f1cdf49-665a-63e9-d221-a1abf756b501
ms.date: 06/08/2017
---


# ColorsInUse.Count Property (Publisher)

Returns a  **Long** that represents the number of items in the specified collection.


## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **ColorsInUse** object.


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


