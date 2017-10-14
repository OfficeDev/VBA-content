---
title: TablesOfFigures Object (Word)
keywords: vbawd10.chm2338
f1_keywords:
- vbawd10.chm2338
ms.prod: word
ms.assetid: 2a5b3c3d-bb25-e31d-e7d3-b011732de6fb
ms.date: 06/08/2017
---


# TablesOfFigures Object (Word)

A collection of  **[TableOfFigures](tableoffigures-object-word.md)** objects that represent the tables of figures in a document.


## Remarks

Use the  **TablesOfFigures** property to return the **TablesOfFigures** collection. The following example applies the Classic format to all tables of figures in the active document.


```vb
ActiveDocument.TablesOfFigures.Format = wdTOFClassic
```

Use the  **Add** method to add a table of figures to a document. A table of figures lists figure captions in the order in which they appear in the document. The following example replaces the selection in the active document with a table of figures that includes caption labels and page numbers.




```vb
ActiveDocument.TablesOfFigures.Add Range:=Selection.Range, _ 
 IncludeLabel:=True, IncludePageNumbers:=True
```

Use  **TablesOfFigures** (Index), where Index is the index number, to return a single **TableOfFigures** object. The index number represents the position of the table of figures in the document. The following example updates the page numbers of the items in the first table of figures in the active document.




```vb
ActiveDocument.TablesOfFigures(1).UpdatePageNumbers
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

