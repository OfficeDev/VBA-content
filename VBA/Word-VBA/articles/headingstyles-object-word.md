---
title: HeadingStyles Object (Word)
ms.prod: word
ms.assetid: be882a12-1264-8f7e-415b-b8bcbf28e703
ms.date: 06/08/2017
---


# HeadingStyles Object (Word)

A collection of  **HeadingStyle** objects that represent the styles used to compile a table of figures or table of contents.


## Remarks

Use the  **HeadingStyles** property to return the **HeadingStyles** collection. The following example displays the number of items in the **HeadingStyles** collection for the first table of contents in the active document.


```vb
MsgBox ActiveDocument.TablesOfContents(1).HeadingStyles.Count
```

Use the  **Add** method to add a style to the **HeadingStyles** collection. The following example adds a table of contents at the beginning of the active document and then adds the Title style to the list of styles used to build a table of contents.




```vb
Set myToc = ActiveDocument.TablesOfContents.Add _ 
 (Range:=ActiveDocument.Range(0, 0), UseHeadingStyles:=True, _ 
 LowerHeadingLevel:=3, UpperHeadingLevel:=1) 
myToc.HeadingStyles.Add Style:="Title", Level:=2
```

Use  **HeadingStyles** (Index), where Index is the index number, to return a single **[HeadingStyle](headingstyle-object-word.md)** object. The index number represents the position of the style in the **HeadingStyles** collection. The following example adds (at the beginning of the active document) a table of figures built from the Title style, and then displays the name of the first style in the **HeadingStyles** collection.




```vb
Set myTOF = ActiveDocument.TablesOfFigures.Add _ 
 (Range:=ActiveDocument.Range(0, 0), AddedStyles:="Title") 
MsgBox myTOF.HeadingStyles(1).Style
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

