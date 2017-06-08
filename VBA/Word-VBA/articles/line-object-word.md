---
title: Line Object (Word)
keywords: vbawd10.chm757
f1_keywords:
- vbawd10.chm757
ms.prod: word
api_name:
- Word.Line
ms.assetid: 1fbc9a15-c677-0f79-4311-9e6de6fc1b27
ms.date: 06/08/2017
---


# Line Object (Word)

Represents an individual line in a **Rectangle** object of type wdTextRectangle. Use the **Line** object and related methods and properties to programmatically define page layout in a document.


## Remarks

Use the  **Item** method to return a specific **Line** object. The following example accesses the first line in the first rectangle in the first page of the active document.


```vb
Dim objLine As Line 
 
Set objLine = ActiveDocument.ActiveWindow _ 
 .Panes(1).Pages(1).Rectangles(1).Lines.Item(1)
```

Use the  **LineType** property to determine whether the specified line is a text line ( **wdTextLine** ) or a table row ( **wdTableRow** ). Then use the **Range** property to access the contents and formatting for the line. The following example creates a reference to the table if the specified line type is **wdTableRow** .




```vb
Dim objLine As Line 
Dim objTable As Table 
 
Set objLine = ActiveDocument.ActiveWindow _ 
 .Panes(1).Pages(1).Rectangles(1).Lines.Item(1) 
 
If objLine.LineType = wdTableRow Then _ 
 Set objTable = objLine.Range.Tables(1)
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

