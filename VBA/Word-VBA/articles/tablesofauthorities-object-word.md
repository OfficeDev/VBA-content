---
title: TablesOfAuthorities Object (Word)
keywords: vbawd10.chm2322
f1_keywords:
- vbawd10.chm2322
ms.prod: word
ms.assetid: c0fd88b1-b737-2811-ec4c-1fc274fc3e20
ms.date: 06/08/2017
---


# TablesOfAuthorities Object (Word)

A collection of  **[TableOfAuthorities](tableofauthorities-object-word.md)** objects (TOA fields) that represents the tables of authorities in a document.


## Remarks

Use the  **TablesOfAuthorities** property to return the **TablesOfAuthorities** collection. The following example applies the Classic built-in format to all the tables of authorities in the active document.


```vb
ActiveDocument.TablesOfAuthorities.Format = wdTOAClassic
```

Use the  **Add** method to add a table of authorities to a document. A table of authorities is built from TA (Table of Authorities Entry) fields in a document. The following example adds a table of authorities that includes all categories at the beginning of the active document.




```vb
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
ActiveDocument.TablesOfAuthorities.Add Range:=myRange, _ 
 Passim:=True, Category:=0, EntrySeparator:= ", "
```

Use  **TablesOfAuthorities** (Index), where Index is the index number, to return a single **TableOfAuthorities** object. The index number represents the position of the table of authorities in the document. The following example includes category headers in the first table of authorities in the active document and then updates the table.




```vb
With ActiveDocument.TablesOfAuthorities(1) 
 .IncludeCategoryHeader = True 
 .Update 
End With
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

