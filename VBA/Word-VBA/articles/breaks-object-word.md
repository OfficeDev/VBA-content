---
title: Breaks Object (Word)
keywords: vbawd10.chm777
f1_keywords:
- vbawd10.chm777
ms.prod: word
api_name:
- Word.Breaks
ms.assetid: c6a9213d-c4bc-a9a6-3889-a32446f1fc25
ms.date: 06/08/2017
---


# Breaks Object (Word)

A collection of page, column, or section breaks in a page. Use the  **Breaks** collection and the related objects and properties to programmatically define page layout in a document.


## Remarks

Use the  **[Breaks](page-breaks-property-word.md)** property to return a **Breaks** collection. The following example returns the breaks in the first page of the active document.


```vb
Dim objBreaks As Breaks 
 
Set objBreaks = ActiveDocument.ActiveWindow _ 
 .Panes(1).Pages(1).Breaks
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

