---
title: TableOfContents Object (Word)
ms.prod: word
api_name:
- Word.TableOfContents
ms.assetid: 629a03c1-ae97-649d-7ec4-25210b4b9ecd
ms.date: 06/08/2017
---


# TableOfContents Object (Word)

Represents a single table of contents in a document. The  **TableOfContents** object is a member of the **[TablesOfContents](tablesofcontents-object-word.md)** collection. The **TablesOfContents** collection includes all the tables of contents in a document.


## Remarks

Use  **TablesOfContents** (Index), where Index is the index number, to return a single **TableOfContents** object. The index number represents the position of the table of contents in the document. The following example updates the page numbers of the items in the first table of figures in the active document.


```
ActiveDocument.TablesOfContents(1).UpdatePageNumbers
```

Use the  **Add** method to add a table of contents to a document. The following example adds a table of contents at the beginning of the active document. The example builds the table of contents from all paragraphs styled as either Heading 1, Heading 2, or Heading 3.




```
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
ActiveDocument.TablesOfContents.Add Range:=myRange, _ 
 UseFields:=False, UseHeadingStyles:=True, _ 
 LowerHeadingLevel:=3, _ 
 UpperHeadingLevel:=1
```


## Methods



|**Name**|
|:-----|
|[Delete](tableofcontents-delete-method-word.md)|
|[Update](tableofcontents-update-method-word.md)|
|[UpdatePageNumbers](tableofcontents-updatepagenumbers-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](tableofcontents-application-property-word.md)|
|[Creator](tableofcontents-creator-property-word.md)|
|[HeadingStyles](tableofcontents-headingstyles-property-word.md)|
|[HidePageNumbersInWeb](tableofcontents-hidepagenumbersinweb-property-word.md)|
|[IncludePageNumbers](tableofcontents-includepagenumbers-property-word.md)|
|[LowerHeadingLevel](tableofcontents-lowerheadinglevel-property-word.md)|
|[Parent](tableofcontents-parent-property-word.md)|
|[Range](tableofcontents-range-property-word.md)|
|[RightAlignPageNumbers](tableofcontents-rightalignpagenumbers-property-word.md)|
|[TabLeader](tableofcontents-tableader-property-word.md)|
|[TableID](tableofcontents-tableid-property-word.md)|
|[UpperHeadingLevel](tableofcontents-upperheadinglevel-property-word.md)|
|[UseFields](tableofcontents-usefields-property-word.md)|
|[UseHeadingStyles](tableofcontents-useheadingstyles-property-word.md)|
|[UseHyperlinks](tableofcontents-usehyperlinks-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
