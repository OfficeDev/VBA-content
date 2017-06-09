---
title: Pages Object (Word)
keywords: vbawd10.chm1026
f1_keywords:
- vbawd10.chm1026
ms.prod: word
ms.assetid: d51e5c61-5719-c70f-b244-99507889f2dc
ms.date: 06/08/2017
---


# Pages Object (Word)

A collection of pages in a document. Use the  **Pages** collection and the related objects and properties for programmatically defining page layout in a document.


## Remarks

Use the  **Pages** property to return a **Pages** collection. The following example accesses all pages in the active document.


```vb
Dim objPages As Pages 
 
Set objPage = ActiveDocument. _ 
 ActiveWindow.Panes(1).Pages
```

Use the  **Item** method to access an individual **Page** object that represents an individual page in a document. The following example accesses the first page in the active document.




```vb
Dim objPage As Page 
 
Set objPage = ActiveDocument.ActiveWindow _ 
 .Panes(1).Pages.Item(1)
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

