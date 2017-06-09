---
title: Break Object (Word)
keywords: vbawd10.chm3057
f1_keywords:
- vbawd10.chm3057
ms.prod: word
api_name:
- Word.Break
ms.assetid: 771ba998-c22e-3cf0-fab7-af9329793855
ms.date: 06/08/2017
---


# Break Object (Word)

Represents individual page, column, and section breaks in a page. Use the  **Break** object and the related methods and properties for programmatically defining page layout in a document.


## Remarks

Use the  **[Item](breaks-item-method-word.md)** method of the **[Breaks](breaks-object-word.md)** collection to return a specific **Break** object. The following example returns the first break in the first page of the active document.


```vb
Dim objBreak As Break 
 
Set objBreak = ActiveDocument.ActiveWindow _ 
 .Panes(1).Pages(1).Breaks.Item(1)
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


