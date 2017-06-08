---
title: Lines Object (Word)
ms.prod: word
api_name:
- Word.Lines
ms.assetid: d04aff17-bd9c-8340-f3ab-191da921ea79
ms.date: 06/08/2017
---


# Lines Object (Word)

A collection of  **Line** objects that represents the lines in a **Rectangle** object that is of type **wdTextRectangle** .


## Remarks

Use the  **Lines** property to return a collection of lines for a specified rectangle. The following example accesses the lines in the first rectangle in the first page in the active document.


```vb
Dim objLines As Lines 
 
Set objLines = ActiveDocument.ActiveWindow.Panes(1) _ 
 .Pages(1).Rectangles(1).Lines
```

Use the  **RectangleType** property of the specified **Rectangle** object to determine whether the **Rectangle** object is of type **wdTextRectangle** . The following example returns the collection of lines in the first rectangle in the first page of the active document if the specified rectangle contains text.




```vb
Dim objRectangle As Rectangle 
Dim objLines As Lines 
 
Set objRectangle = ActiveDocument.ActiveWindow _ 
 .Panes(1).Pages(1).Rectangles(1) 
 
If objRectangle.RectangleType = wdTextRectangle Then _ 
 Set objLines = objRectangle.Lines
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

