---
title: TablesOfContents Object (Word)
keywords: vbawd10.chm2324
f1_keywords:
- vbawd10.chm2324
ms.prod: word
ms.assetid: d0d0e5fc-e443-31ae-e1a9-15b945f1e318
ms.date: 06/08/2017
---


# TablesOfContents Object (Word)

A collection of  **[TableOfContents](tableofcontents-object-word.md)** objects that represent the tables of contents in a document.


## Remarks

Use the  **TablesOfContents** property to return the **TablesOfContents** collection. The following example inserts a table of contents entry that references the selected text in the active document.


```
ActiveDocument.TablesOfContents.MarkEntry Range:=Selection.Range, _ 
 Level:=2, Entry:="Introduction"
```

Use the  **Add** method to add a table of contents to a document. The following example adds a table of contents at the beginning of the active document. The example builds the table of contents from all paragraphs styled as either Heading 1, Heading 2, or Heading 3.




```
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
ActiveDocument.TablesOfContents.Add Range:=myRange, _ 
 UseFields:=False, UseHeadingStyles:=True, _ 
 LowerHeadingLevel:=3, _ 
 UpperHeadingLevel:=1
```

Use  **TablesOfContents** (Index), where Index is the index number, to return a single **TableOfContents** object. The index number represents the position of the table of contents in the document. The following example updates the page numbers of the items in the first table of figures in the active document.




```
ActiveDocument.TablesOfContents(1).UpdatePageNumbers
```


## Methods



|**Name**|
|:-----|
|[Add](tablesofcontents-add-method-word.md)|
|[Item](tablesofcontents-item-method-word.md)|
|[MarkEntry](tablesofcontents-markentry-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](tablesofcontents-application-property-word.md)|
|[Count](tablesofcontents-count-property-word.md)|
|[Creator](tablesofcontents-creator-property-word.md)|
|[Format](tablesofcontents-format-property-word.md)|
|[Parent](tablesofcontents-parent-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
